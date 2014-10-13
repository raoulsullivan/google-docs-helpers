#TODO: sort logging, better worksheet identifies (helpers?)

import gdata.spreadsheet
import gdata.docs.service
import gdata.spreadsheet.service
import re
from logging_setup import logFactory
from collections import OrderedDict

class GoogleSpreadsheetsClient():
    """ Set up Google spreadsheets client """

    def __init__(self, gmailUsername, gmailPassword, sSourceName='Default', bDebug=False):
        self.log = logFactory('googleSpreadsheetsLogger',debug=bDebug)
        self.spr_client = gdata.spreadsheet.service.SpreadsheetsService()
        self.spr_client.email = gmailUsername
        self.spr_client.password = gmailPassword
        self.spr_client.source = sSourceName
        self.spr_client.debug = bDebug
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

    def EscapeHeader(self,header):
        """Makes a header that gdocs can deal with """
        header = header.lower()
        #header = re.sub('[\s_]+','-',header)
        header = re.sub('[^0-9a-z-]+','',header)
        return(header)

    def CreateTable(self,sSpreadsheetKey,sWorksheetId,aRows):
        """ Creates a new table from scratch """
        self.log.info('Creating a new table on worksheet %s',sWorksheetId)
        aHeaders = []
        for key, value in aRows[0].items():
            aHeaders.append(self.EscapeHeader(key))
        self.CreateTableHeaders(sSpreadsheetKey,sWorksheetId,aHeaders)
        #ok, for some reason it seems to throw errors on some sheets when trying to add
        #a row at the bottom where there isn't space yet.
        #so let's add in all the rows in advance
        sheets = self.spr_client.GetWorksheetsFeed(sSpreadsheetKey)
        for sheet in sheets.entry:
            if sheet._GDataEntry__id.text.split('/')[-1] == sWorksheetId:
                targetsheet=sheet

        if targetsheet:
            targetsheet.row_count.text = str(len(aRows)+1)
            self.spr_client.UpdateWorksheet(targetsheet)

        #now add the rows
        i = 0
        for row in aRows:
            dRow = {}
            for key, value in row.items():
                dRow[self.EscapeHeader(key)] = str(value)
            self.log.debug(dRow)
            errorcount = 0
            maxerror = 3
            while errorcount < maxerror:
                try:
                    self.spr_client.InsertRow(dRow,sSpreadsheetKey,wksht_id=sWorksheetId)
                except Exception as e:
                    self.log.warning(e)
                    obj = json.loads(e.msg)
                    if not (obj['status'] in [500,502]):
                        raise
                    errorcount += 1
                    self.log.warning('Retrying - attempt {0} of {1}'.format(errorcount,maxerror))
                    continue
                i += 1
                break

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
        #list feed method fails to clear empty rows
        #lfeed = gdata.spreadsheet.GetListFeed(key=sSpreadsheetKey,wksht_id=sWorksheetId)
        #for row in lfeed.entry.reverse():
        #    gdata.spreadsheet.DeleteRow(row)

        sheets = self.spr_client.GetWorksheetsFeed(sSpreadsheetKey)
        for sheet in sheets.entry:
            if sheet._GDataEntry__id.text.split('/')[-1] == sWorksheetId:
                targetsheet=sheet

        if targetsheet:
            targetsheet.row_count.text = '1'
            self.spr_client.UpdateWorksheet(targetsheet)

        #This method parked as it doesn't seem to actually delete the rows?
        #Can be used to clear the header, though...
        batch = gdata.spreadsheet.SpreadsheetsCellsFeed()
        fWorksheetCellsFeed = self.spr_client.GetCellsFeed(key=sSpreadsheetKey,wksht_id=sWorksheetId)
        for i, entry in enumerate(fWorksheetCellsFeed.entry):
            entry.cell.inputValue = ''
            batch.AddUpdate(fWorksheetCellsFeed.entry[i])

        self.spr_client.ExecuteBatch(batch, fWorksheetCellsFeed.GetBatchLink().href)

    def GetHeadersFromWorksheet(self,sSpreadsheetKey,sWorksheetId):
        """Gets the headers from a worksheet"""
        headers = []
        feed = self.spr_client.GetCellsFeed(sSpreadsheetKey,sWorksheetId)
        for i, entry in enumerate(feed.entry):
        #iterate through the cells feed until the second row, taking the cell contents as headers
            if int(entry.cell.row) >= 2:
                break
            headers.append(self.EscapeHeader(entry.content.text))
        return(headers)

    def GetRowsFromWorksheet(self,sSpreadsheetKey,sWorksheetId):
        """Gets the rows from a worksheet"""
        headers = self.GetHeadersFromWorksheet(sSpreadsheetKey,sWorksheetId)
        rows = []
        feed = self.spr_client.GetListFeed(key=sSpreadsheetKey,wksht_id=sWorksheetId)
        for i, entry in enumerate(feed.entry):
            row = OrderedDict()
            row['id'] = i
            for header in headers:
                row[header] = entry.custom[header].text
            rows.append(row)
        return(rows)

    def PutRowsIntoWorksheet(self,sSpreadsheetKey,sWorksheetId,aRows):
        """Appeands a row to a worksheet"""
        for dRow in aRows:
            entry = self.spr_client.InsertRow(dRow,sSpreadsheetKey,sWorksheetId)
