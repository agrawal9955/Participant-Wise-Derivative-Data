import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread_formatting import format_cell_range, cellFormat, color, textFormat
from GlobalConstants import GlobalConstants
from NseDataModel import NseHistoricalData, ParticipantWiseNseRawRecords


class GoogleSheetService:

    def __init__(self):
        # Authorize the API
        self.globalConstants = GlobalConstants()
        self.creds = ServiceAccountCredentials.from_json_keyfile_name(
            "fin-position-excel-service-60f7d648377e.json",
            [
                'https://www.googleapis.com/auth/drive',
                'https://www.googleapis.com/auth/drive.file'
            ])
        client = gspread.authorize(self.creds)
        # Fetch the sheet
        self.ss = client.open(self.globalConstants.getGoogleWorksheetName())
        worksheets = self.ss.worksheets()
        self.formattedDataSheet = worksheets[self.globalConstants.getFormattedDataSheet(
        )]
        self.formattedDataSheetId = self.formattedDataSheet._properties['sheetId']
        self.rawDataSheet = worksheets[self.globalConstants.getRawDataSheet(
        )]
        # contains number of rows consumed by raw data sheet
        self.pythonScriptMetaData = worksheets[self.globalConstants.getMetaDataSheet(
        )]
        self.formattedFlatDataSheet = worksheets[self.globalConstants.getFormattedFlatDataSheet(
        )]

    def addParticipantRawData(self, data: ParticipantWiseNseRawRecords):
        self.rawDataSheet.append_row(data.dateHeader)
        self.rawDataSheet.append_row(data.columnHeader)
        self.rawDataSheet.append_row(data.clientData)
        self.rawDataSheet.append_row(data.diiData)
        self.rawDataSheet.append_row(data.fiiData)
        self.rawDataSheet.append_row(data.proData)
        self.rawDataSheet.append_row(data.totalData)
        time.sleep(30)

        lastRowCount = self.getLastRowRawData()
        self.pythonScriptMetaData.update_cell(1, 1, lastRowCount + 7)
        return

    def addParticipantFormattedFlatData(self, date, niftyData: NseHistoricalData):
        lastDataRowCount = self.getLastRowFormattedData()
        lastRowCount = self.getLastRowFormattedFlatData()
        self.formattedFlatDataSheet.append_row([" "])

        # Date
        self.formattedFlatDataSheet.update_cell(
            lastRowCount, 1, date)

        self.formattedFlatDataSheet.update_cell(
            lastRowCount, 2, self.formattedDataSheet.cell(
                float(lastDataRowCount-23), 10).value)

        self.formattedFlatDataSheet.update_cell(
            lastRowCount, 3, self.formattedDataSheet.cell(
                float(lastDataRowCount-22), 10).value)

        self.formattedFlatDataSheet.update_cell(
            lastRowCount, 4, self.formattedDataSheet.cell(
                float(lastDataRowCount-21), 10).value)

        self.formattedFlatDataSheet.update_cell(
            lastRowCount, 5, self.formattedDataSheet.cell(
                float(lastDataRowCount-20), 10).value)

        self.formattedFlatDataSheet.update_cell(
            lastRowCount, 6, self.formattedDataSheet.cell(
                float(lastDataRowCount-17), 10).value)

        self.formattedFlatDataSheet.update_cell(
            lastRowCount, 7, self.formattedDataSheet.cell(
                float(lastDataRowCount-16), 10).value)

        self.formattedFlatDataSheet.update_cell(
            lastRowCount, 8, self.formattedDataSheet.cell(
                float(lastDataRowCount-15), 10).value)

        self.formattedFlatDataSheet.update_cell(
            lastRowCount, 9, self.formattedDataSheet.cell(
                float(lastDataRowCount-14), 10).value)

        self.formattedFlatDataSheet.update_cell(
            lastRowCount, 10, self.formattedDataSheet.cell(
                float(lastDataRowCount-11), 10).value)

        self.formattedFlatDataSheet.update_cell(
            lastRowCount, 11, self.formattedDataSheet.cell(
                float(lastDataRowCount-10), 10).value)

        self.formattedFlatDataSheet.update_cell(
            lastRowCount, 12, self.formattedDataSheet.cell(
                float(lastDataRowCount-9), 10).value)

        self.formattedFlatDataSheet.update_cell(
            lastRowCount, 13, self.formattedDataSheet.cell(
                float(lastDataRowCount-8), 10).value)

        self.formattedFlatDataSheet.update_cell(
            lastRowCount, 14, self.formattedDataSheet.cell(
                float(lastDataRowCount-5), 10).value)

        self.formattedFlatDataSheet.update_cell(
            lastRowCount, 15, self.formattedDataSheet.cell(
                float(lastDataRowCount-4), 10).value)

        self.formattedFlatDataSheet.update_cell(
            lastRowCount, 16, self.formattedDataSheet.cell(
                float(lastDataRowCount-3), 10).value)

        self.formattedFlatDataSheet.update_cell(
            lastRowCount, 17, self.formattedDataSheet.cell(
                float(lastDataRowCount-2), 10).value)

        # NIFTY PriceData
        self.formattedFlatDataSheet.update_cell(
            lastRowCount, 18, niftyData.open)

        self.formattedFlatDataSheet.update_cell(
            lastRowCount, 19, niftyData.close)

        self.formattedFlatDataSheet.update_cell(
            lastRowCount, 20, niftyData.high)

        self.formattedFlatDataSheet.update_cell(
            lastRowCount, 21, niftyData.low)

        # %change
        change = (float(self.formattedFlatDataSheet.cell(
            lastRowCount, 19).value.replace(",", "")) - float(self.formattedFlatDataSheet.cell(
                lastRowCount-1, 19).value.replace(",", "")))*100/float(self.formattedFlatDataSheet.cell(
                    lastRowCount-1, 19).value.replace(",", ""))

        self.formattedFlatDataSheet.update_cell(
            lastRowCount, 22, change)

        self.pythonScriptMetaData.update_cell(2, 1, lastRowCount + 1)
        time.sleep(6)
        return

    def addParticipantDataFormula(self):
        # get last Row Raw Data
        lastRowCountRawData = self.getLastRowRawData()
        currentRawDataRowStart = lastRowCountRawData - 5
        prevRawDataRowStart = lastRowCountRawData - 12

        # get last Row Formatted Data
        lastRowCountFormattedData = self.getLastRowFormattedData()

        for index in range(0, 5):
            self.formattedDataSheet.append_row(
                [" ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", "."])
            lastRowCountFormattedData = lastRowCountFormattedData + 1

        self.formattedDataSheet.append_row([" "])
        self.formattedDataSheet.update_cell(
            lastRowCountFormattedData, 1, "='Daily OI Data - Raw Data'!E" + str(lastRowCountRawData - 7))

        lastRowCountFormattedData = self.__updateSheetWithFormattedData(
            lastRowCountFormattedData, currentRawDataRowStart, prevRawDataRowStart)

        self.__mergeDateCol(lastRowCountFormattedData)

        self.pythonScriptMetaData.update_cell(1, 2, lastRowCountFormattedData)

        return

    def getLastRowRawData(self):
        return int(self.pythonScriptMetaData.cell(1, 1).value)

    def getLastRowFormattedFlatData(self):
        return int(self.pythonScriptMetaData.cell(2, 1).value)

    def getLastRowFormattedData(self):
        return int(self.pythonScriptMetaData.cell(1, 2).value)

    def getLastDate(self) -> str:
        return str(self.pythonScriptMetaData.cell(1, 3).value)

    def setLastDate(self, date) -> str:
        self.pythonScriptMetaData.update_cell(1, 3, date)

    def __updateSheetWithFormattedData(self, currentRowCount, currentRawDataRowStart, prevRawDataRowStart):
        # Add Col Headers
        self.formattedDataSheet.append_row([" "])
        self.formattedDataSheet.update_cell(
            currentRowCount, 3, "Position Change")
        self.formattedDataSheet.update_cell(currentRowCount, 7, "Net Position")
        self.formattedDataSheet.update_cell(currentRowCount, 10, "Net Change")

        updatedRowCount = currentRowCount + 1

        updatedRowCount = self.__updateSheetWithIndexFuturesData(
            updatedRowCount, currentRawDataRowStart, prevRawDataRowStart)
        updatedRowCount = self.__updateSheetWithIndexCallsData(
            updatedRowCount, currentRawDataRowStart, prevRawDataRowStart)
        updatedRowCount = self.__updateSheetWithIndexPutsData(
            updatedRowCount, currentRawDataRowStart, prevRawDataRowStart)
        updatedRowCount = self.__updateSheetWithStockFuturesData(
            updatedRowCount, currentRawDataRowStart, prevRawDataRowStart)

        return updatedRowCount

    def __updateSheetWithIndexFuturesData(self, currentFormattedRowCount, currentRawDataRowStart, prevRawDataRowStart):
        # Add Col Headers
        self.formattedDataSheet.append_row([" "])
        self.formattedDataSheet.update_cell(
            currentFormattedRowCount, 2, "Index Futures")
        self.formattedDataSheet.update_cell(
            currentFormattedRowCount, 3, "Longs")
        self.formattedDataSheet.update_cell(
            currentFormattedRowCount, 4, "% Change")
        self.formattedDataSheet.update_cell(
            currentFormattedRowCount, 5, "Shorts")
        self.formattedDataSheet.update_cell(
            currentFormattedRowCount, 6, "% Change")
        self.formattedDataSheet.update_cell(
            currentFormattedRowCount, 7, "Today")
        self.formattedDataSheet.update_cell(
            currentFormattedRowCount, 8, "1 Day Ago")
        self.formattedDataSheet.update_cell(
            currentFormattedRowCount, 9, "2 Day Ago")
        self.formattedDataSheet.update_cell(
            currentFormattedRowCount, 10, "Net Change")
        currentFormattedRowCount = currentFormattedRowCount + 1
        time.sleep(6)

        self.formattedDataSheet.append_rows(
            [[" "], [" "], [" "], [" "], [" "]])
        # Clients Data
        for index in range(0, 5):
            # self.formattedDataSheet.append_row([" "])
            self.formattedDataSheet.update_cell(
                currentFormattedRowCount, 2, "='Daily OI Data - Raw Data'!A" + str(currentRawDataRowStart + index))
            self.formattedDataSheet.update_cell(currentFormattedRowCount, 3, "='Daily OI Data - Raw Data'!B" + str(
                currentRawDataRowStart + index) + "- 'Daily OI Data - Raw Data'!B" + str(prevRawDataRowStart + index))
            self.formattedDataSheet.update_cell(currentFormattedRowCount, 4, "=( 'Daily OI Data - Raw Data'!B" + str(currentRawDataRowStart + index) +
                                                "- 'Daily OI Data - Raw Data'!B" + str(prevRawDataRowStart + index) + " ) * 100 / " + "'Daily OI Data - Raw Data'!B" + str(prevRawDataRowStart + index))
            self.formattedDataSheet.update_cell(currentFormattedRowCount, 5, "='Daily OI Data - Raw Data'!C" + str(
                currentRawDataRowStart + index) + "- 'Daily OI Data - Raw Data'!C" + str(prevRawDataRowStart + index))
            self.formattedDataSheet.update_cell(currentFormattedRowCount, 6, "=( 'Daily OI Data - Raw Data'!C" + str(currentRawDataRowStart + index) +
                                                "- 'Daily OI Data - Raw Data'!C" + str(prevRawDataRowStart + index) + " ) * 100 / " + "'Daily OI Data - Raw Data'!C" + str(prevRawDataRowStart + index))
            self.formattedDataSheet.update_cell(currentFormattedRowCount, 7, "='Daily OI Data - Raw Data'!B" + str(
                currentRawDataRowStart + index) + "- 'Daily OI Data - Raw Data'!C" + str(currentRawDataRowStart + index))
            self.formattedDataSheet.update_cell(
                currentFormattedRowCount, 8, "='Daily OI Analysis - Formatted Data - 2'!G" + str(currentFormattedRowCount - 30))
            self.formattedDataSheet.update_cell(
                currentFormattedRowCount, 9, "='Daily OI Analysis - Formatted Data - 2'!H" + str(currentFormattedRowCount - 30))
            self.formattedDataSheet.update_cell(currentFormattedRowCount, 10, "='Daily OI Analysis - Formatted Data - 2'!G" + str(
                currentFormattedRowCount) + " - 'Daily OI Analysis - Formatted Data - 2'!H" + str(currentFormattedRowCount))
            currentFormattedRowCount = currentFormattedRowCount + 1
            time.sleep(6)

        return currentFormattedRowCount

    def __updateSheetWithIndexCallsData(self, currentFormattedRowCount, currentRawDataRowStart, prevRawDataRowStart):
        # Add Col Headers
        self.formattedDataSheet.append_row([" "])
        self.formattedDataSheet.update_cell(
            currentFormattedRowCount, 2, "Index Calls")
        self.formattedDataSheet.update_cell(
            currentFormattedRowCount, 3, "Longs")
        self.formattedDataSheet.update_cell(
            currentFormattedRowCount, 4, "% Change")
        self.formattedDataSheet.update_cell(
            currentFormattedRowCount, 5, "Shorts")
        self.formattedDataSheet.update_cell(
            currentFormattedRowCount, 6, "% Change")
        self.formattedDataSheet.update_cell(
            currentFormattedRowCount, 7, "Today")
        self.formattedDataSheet.update_cell(
            currentFormattedRowCount, 8, "1 Day Ago")
        self.formattedDataSheet.update_cell(
            currentFormattedRowCount, 9, "2 Day Ago")
        self.formattedDataSheet.update_cell(
            currentFormattedRowCount, 10, "Net Change")
        currentFormattedRowCount = currentFormattedRowCount + 1
        time.sleep(6)

        self.formattedDataSheet.append_rows(
            [[" "], [" "], [" "], [" "], [" "]])

        # Clients Data
        for index in range(0, 5):
            self.formattedDataSheet.update_cell(
                currentFormattedRowCount, 2, "='Daily OI Data - Raw Data'!A" + str(currentRawDataRowStart + index))
            self.formattedDataSheet.update_cell(currentFormattedRowCount, 3, "='Daily OI Data - Raw Data'!F" + str(
                currentRawDataRowStart + index) + "- 'Daily OI Data - Raw Data'!F" + str(prevRawDataRowStart + index))
            self.formattedDataSheet.update_cell(currentFormattedRowCount, 4, "=( 'Daily OI Data - Raw Data'!F" + str(currentRawDataRowStart + index) +
                                                "- 'Daily OI Data - Raw Data'!F" + str(prevRawDataRowStart + index) + " ) * 100 / " + "'Daily OI Data - Raw Data'!F" + str(prevRawDataRowStart + index))
            self.formattedDataSheet.update_cell(currentFormattedRowCount, 5, "='Daily OI Data - Raw Data'!H" + str(
                currentRawDataRowStart + index) + "- 'Daily OI Data - Raw Data'!H" + str(prevRawDataRowStart + index))
            self.formattedDataSheet.update_cell(currentFormattedRowCount, 6, "=( 'Daily OI Data - Raw Data'!H" + str(currentRawDataRowStart + index) +
                                                "- 'Daily OI Data - Raw Data'!H" + str(prevRawDataRowStart + index) + " ) * 100 / " + "'Daily OI Data - Raw Data'!H" + str(prevRawDataRowStart + index))
            self.formattedDataSheet.update_cell(currentFormattedRowCount, 7, "='Daily OI Data - Raw Data'!F" + str(
                currentRawDataRowStart + index) + "- 'Daily OI Data - Raw Data'!H" + str(currentRawDataRowStart + index))
            self.formattedDataSheet.update_cell(
                currentFormattedRowCount, 8, "='Daily OI Analysis - Formatted Data - 2'!G" + str(currentFormattedRowCount - 30))
            self.formattedDataSheet.update_cell(
                currentFormattedRowCount, 9, "='Daily OI Analysis - Formatted Data - 2'!H" + str(currentFormattedRowCount - 30))
            self.formattedDataSheet.update_cell(currentFormattedRowCount, 10, "='Daily OI Analysis - Formatted Data - 2'!G" + str(
                currentFormattedRowCount) + " - 'Daily OI Analysis - Formatted Data - 2'!H" + str(currentFormattedRowCount))
            currentFormattedRowCount = currentFormattedRowCount + 1
            time.sleep(6)

        return currentFormattedRowCount

    def __updateSheetWithIndexPutsData(self, currentFormattedRowCount, currentRawDataRowStart, prevRawDataRowStart):
        # Add Col Headers
        self.formattedDataSheet.append_row([" "])
        self.formattedDataSheet.update_cell(
            currentFormattedRowCount, 2, "Index Puts")
        self.formattedDataSheet.update_cell(
            currentFormattedRowCount, 3, "Longs")
        self.formattedDataSheet.update_cell(
            currentFormattedRowCount, 4, "% Change")
        self.formattedDataSheet.update_cell(
            currentFormattedRowCount, 5, "Shorts")
        self.formattedDataSheet.update_cell(
            currentFormattedRowCount, 6, "% Change")
        self.formattedDataSheet.update_cell(
            currentFormattedRowCount, 7, "Today")
        self.formattedDataSheet.update_cell(
            currentFormattedRowCount, 8, "1 Day Ago")
        self.formattedDataSheet.update_cell(
            currentFormattedRowCount, 9, "2 Day Ago")
        self.formattedDataSheet.update_cell(
            currentFormattedRowCount, 10, "Net Change")
        currentFormattedRowCount = currentFormattedRowCount + 1
        time.sleep(6)

        self.formattedDataSheet.append_rows(
            [[" "], [" "], [" "], [" "], [" "]])

        # Clients Data
        for index in range(0, 5):
            self.formattedDataSheet.update_cell(
                currentFormattedRowCount, 2, "='Daily OI Data - Raw Data'!A" + str(currentRawDataRowStart + index))
            self.formattedDataSheet.update_cell(currentFormattedRowCount, 3, "='Daily OI Data - Raw Data'!G" + str(
                currentRawDataRowStart + index) + "- 'Daily OI Data - Raw Data'!G" + str(prevRawDataRowStart + index))
            self.formattedDataSheet.update_cell(currentFormattedRowCount, 4, "=( 'Daily OI Data - Raw Data'!G" + str(currentRawDataRowStart + index) +
                                                "- 'Daily OI Data - Raw Data'!G" + str(prevRawDataRowStart + index) + " ) * 100 / " + "'Daily OI Data - Raw Data'!G" + str(prevRawDataRowStart + index))
            self.formattedDataSheet.update_cell(currentFormattedRowCount, 5, "='Daily OI Data - Raw Data'!I" + str(
                currentRawDataRowStart + index) + "- 'Daily OI Data - Raw Data'!I" + str(prevRawDataRowStart + index))
            self.formattedDataSheet.update_cell(currentFormattedRowCount, 6, "=( 'Daily OI Data - Raw Data'!I" + str(currentRawDataRowStart + index) +
                                                "- 'Daily OI Data - Raw Data'!I" + str(prevRawDataRowStart + index) + " ) * 100 / " + "'Daily OI Data - Raw Data'!I" + str(prevRawDataRowStart + index))
            self.formattedDataSheet.update_cell(currentFormattedRowCount, 7, "='Daily OI Data - Raw Data'!G" + str(
                currentRawDataRowStart + index) + "- 'Daily OI Data - Raw Data'!I" + str(currentRawDataRowStart + index))
            self.formattedDataSheet.update_cell(
                currentFormattedRowCount, 8, "='Daily OI Analysis - Formatted Data - 2'!G" + str(currentFormattedRowCount - 30))
            self.formattedDataSheet.update_cell(
                currentFormattedRowCount, 9, "='Daily OI Analysis - Formatted Data - 2'!H" + str(currentFormattedRowCount - 30))
            self.formattedDataSheet.update_cell(currentFormattedRowCount, 10, "='Daily OI Analysis - Formatted Data - 2'!G" + str(
                currentFormattedRowCount) + " - 'Daily OI Analysis - Formatted Data - 2'!H" + str(currentFormattedRowCount))
            currentFormattedRowCount = currentFormattedRowCount + 1
            time.sleep(6)

        return currentFormattedRowCount

    def __updateSheetWithStockFuturesData(self, currentFormattedRowCount, currentRawDataRowStart, prevRawDataRowStart):
        # Add Col Headers
        self.formattedDataSheet.append_row([" "])
        self.formattedDataSheet.update_cell(
            currentFormattedRowCount, 2, "Stock Futures")
        self.formattedDataSheet.update_cell(
            currentFormattedRowCount, 3, "Longs")
        self.formattedDataSheet.update_cell(
            currentFormattedRowCount, 4, "% Change")
        self.formattedDataSheet.update_cell(
            currentFormattedRowCount, 5, "Shorts")
        self.formattedDataSheet.update_cell(
            currentFormattedRowCount, 6, "% Change")
        self.formattedDataSheet.update_cell(
            currentFormattedRowCount, 7, "Today")
        self.formattedDataSheet.update_cell(
            currentFormattedRowCount, 8, "1 Day Ago")
        self.formattedDataSheet.update_cell(
            currentFormattedRowCount, 9, "2 Day Ago")
        self.formattedDataSheet.update_cell(
            currentFormattedRowCount, 10, "Net Change")
        currentFormattedRowCount = currentFormattedRowCount + 1
        time.sleep(6)

        self.formattedDataSheet.append_rows(
            [[" "], [" "], [" "], [" "], [" "]])

        # Clients Data
        for index in range(0, 5):
            self.formattedDataSheet.update_cell(
                currentFormattedRowCount, 2, "='Daily OI Data - Raw Data'!A" + str(currentRawDataRowStart + index))
            self.formattedDataSheet.update_cell(currentFormattedRowCount, 3, "='Daily OI Data - Raw Data'!D" + str(
                currentRawDataRowStart + index) + "- 'Daily OI Data - Raw Data'!D" + str(prevRawDataRowStart + index))
            self.formattedDataSheet.update_cell(currentFormattedRowCount, 4, "=('Daily OI Data - Raw Data'!D" + str(currentRawDataRowStart + index) +
                                                "- 'Daily OI Data - Raw Data'!D" + str(prevRawDataRowStart + index) + " ) * 100 / " + "'Daily OI Data - Raw Data'!D" + str(prevRawDataRowStart + index))
            self.formattedDataSheet.update_cell(currentFormattedRowCount, 5, "='Daily OI Data - Raw Data'!E" + str(
                currentRawDataRowStart + index) + "- 'Daily OI Data - Raw Data'!E" + str(prevRawDataRowStart + index))
            self.formattedDataSheet.update_cell(currentFormattedRowCount, 6, "=( 'Daily OI Data - Raw Data'!E" + str(currentRawDataRowStart + index) +
                                                "- 'Daily OI Data - Raw Data'!E" + str(prevRawDataRowStart + index) + " ) * 100 / " + "'Daily OI Data - Raw Data'!E" + str(prevRawDataRowStart + index))
            self.formattedDataSheet.update_cell(currentFormattedRowCount, 7, "='Daily OI Data - Raw Data'!D" + str(
                currentRawDataRowStart + index) + "- 'Daily OI Data - Raw Data'!E" + str(currentRawDataRowStart + index))
            self.formattedDataSheet.update_cell(
                currentFormattedRowCount, 8, "='Daily OI Analysis - Formatted Data - 2'!G" + str(currentFormattedRowCount - 30))
            self.formattedDataSheet.update_cell(
                currentFormattedRowCount, 9, "='Daily OI Analysis - Formatted Data - 2'!H" + str(currentFormattedRowCount - 30))
            self.formattedDataSheet.update_cell(currentFormattedRowCount, 10, "='Daily OI Analysis - Formatted Data - 2'!G" + str(
                currentFormattedRowCount) + " - 'Daily OI Analysis - Formatted Data - 2'!H" + str(currentFormattedRowCount))
            currentFormattedRowCount = currentFormattedRowCount + 1
            time.sleep(6)

        return currentFormattedRowCount

    def __mergeDateCol(self, lastFormattedRowCount):
        body = {
            "requests": [
                {
                    "mergeCells": {
                        "mergeType": "MERGE_ALL",
                        "range": {  # In this sample script, all cells of "A1:C3" of "Sheet1" are merged.
                            "sheetId": self.formattedDataSheetId,
                            "startRowIndex": lastFormattedRowCount - 26,
                            "endRowIndex": lastFormattedRowCount - 1,
                            "startColumnIndex": 0,
                            "endColumnIndex": 1
                        }
                    }
                }
            ]
        }
        res = self.ss.batch_update(body)
        return
