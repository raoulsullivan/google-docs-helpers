import gdata.spreadsheet
import gdata.docs.service
import gdata.spreadsheet.service
import re

class GoogleSpreadsheetsClient():
    """ Set up Google spreadsheets client """

    def __init__(self, gmail_username, gmail_password, log, sSourceName='Default'):
        self.log = log
        self.spr_client = gdata.spreadsheet.service.SpreadsheetsService()
        self.spr_client.email = gmail_username
        self.spr_client.password = gmail_password
        self.spr_client.source = sSourceName
        self.spr_client.http_client.debug = True
        self.spr_client.ProgrammaticLogin()
    
    def ExposeClient(self):
        """ Exposes the Google spreadsheets client in case you need to do anything further to it """
        return self.spr_client
    
    def CreateTableHeaders(self,sSpreadsheetKey,sWorksheetId,aHeaders):
        """ Adds table header row """
        self.log.info('Adding headers to worksheet '+sWorksheetId)
        self.log.debug(aHeaders)
        for i, header in enumerate(aHeaders):
            i = i+1
            self.log.debug('%s - %s' % (i, header))
            self.spr_client.UpdateCell(row=1, col=i, inputValue=header, key=sSpreadsheetKey, wksht_id=sWorksheetId)
        self.log.info('Headers added ok')

    def SortHeader(self,header):
        """Makes a header that gdocs can deal with """
        header = header.lower()
        header = re.sub('[\s_]+','-',header)
        header = re.sub('[^0-9a-z-]+','',header)
        return(header)

    def CreateTable(self,sSpreadsheetKey,sWorksheetId,aRows):
        """ Creates a new table from scratch """
        self.log.info('Creating a new table on worksheet %s',sWorksheetId)
        aHeaders = []
        for key, value in aRows[0].items():
            aHeaders.append(self.SortHeader(key))
        self.CreateTableHeaders(sSpreadsheetKey,sWorksheetId,aHeaders)
        i = 0 
        for row in aRows:
            dRow = {}
            for key, value in row.items():
                dRow[self.SortHeader(key)] = str(value)
            self.log.debug(dRow)
            self.spr_client.InsertRow(dRow,sSpreadsheetKey,wksht_id=sWorksheetId)
            i += 1

        self.log.info('%s rows added',i)

    def GetGoogleWorksheets(self,sSpreadsheetKey):
        """ Gets the tabs of the spreadsheet """
        self.log.info('Getting worksheets')
        fWorksheets = self.spr_client.GetWorksheetsFeed(sSpreadsheetKey)
        dWorksheets = {}
        dWorksheets2 = {}
        for i, entry in enumerate(fWorksheets.entry):
            wid = entry.id.text.split('/')[-1]
            name = entry.title.text
            dWorksheets[wid] = name
            dWorksheets2[name] = wid
        ret = {'worksheets_by_id': dWorksheets,'worksheets_by_name': dWorksheets2}
        self.log.info(ret)
        return ret

    def EmptyGoogleWorksheet(self,sSpreadsheetKey,sWorksheetId):
        """ Nukes the worksheet in question from orbit """
        self.log.info('Emptying worksheet ' + sWorksheetId)
        batch = gdata.spreadsheet.SpreadsheetsCellsFeed()
        fWorksheetCellsFeed = self.spr_client.GetCellsFeed(key=sSpreadsheetKey,wksht_id=sWorksheetId)
        for i, entry in enumerate(fWorksheetCellsFeed.entry):
            entry.cell.inputValue = ''
            batch.AddUpdate(fWorksheetCellsFeed.entry[i])

        self.spr_client.ExecuteBatch(batch, fWorksheetCellsFeed.GetBatchLink().href)
