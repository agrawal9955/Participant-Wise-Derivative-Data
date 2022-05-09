import time
import xlwt
import xlrd
from NseDataModel import ParticipantWiseNseRawRecords

class ExcelSheetService:

    def __init__(self):
  
        self.workbook = xlwt.Workbook()

        self.formattedDataSheet = self.workbook.add_sheet("Daily OI analysis")
        # self.rawDataSheet = self.workbook.add_sheet("Daily OI data")
        self.lastRawDataRow = 14
        self.lastFormattedDataRow = 0


        # Give the location of the file
        loc = ("raw_data_nse.xls")

        # To open Workbook
        wb = xlrd.open_workbook(loc)
        self.rawDataSheet = wb.sheet_by_index(1)

    def saveExcel(self):
        self.workbook.save('xlwt example.xls')

    def addParticipantRawData(self, data: ParticipantWiseNseRawRecords):
        # self.lastRawDataRow = self.__appendRow(self.rawDataSheet, self.lastRawDataRow, data.dateHeader)
        # self.lastRawDataRow = self.__appendRow(self.rawDataSheet, self.lastRawDataRow, data.columnHeader)
        # self.lastRawDataRow = self.__appendRow(self.rawDataSheet, self.lastRawDataRow, data.clientData)
        # self.lastRawDataRow = self.__appendRow(self.rawDataSheet, self.lastRawDataRow, data.diiData)
        # self.lastRawDataRow = self.__appendRow(self.rawDataSheet, self.lastRawDataRow, data.fiiData)
        # self.lastRawDataRow = self.__appendRow(self.rawDataSheet, self.lastRawDataRow, data.proData)
        # self.lastRawDataRow = self.__appendRow(self.rawDataSheet, self.lastRawDataRow, data.totalData)

        self.lastRawDataRow = self.lastRawDataRow + 7
        return

    def __appendRow(self, sheet, rowIndex, rowData):
        index = 0
        for data in rowData:
            sheet.write(rowIndex, index, data)
            index = index + 1

        return rowIndex + 1

    def addParticipantFormattedData(self):

        while self.lastRawDataRow < 665: 
            self.addParticipantRawData(None)

            # get last Row Raw Data
            lastRowCountRawData = self.__getLastRowRawData()

            # get last Row Formatted Data
            lastRowCountFormattedData = self.__getLastRowFormattedData()

            self.formattedDataSheet.write(lastRowCountFormattedData, 0, self.rawDataSheet.cell_value(lastRowCountRawData - 7, 4))

            lastRowCountFormattedData = self.__updateSheetWithFormattedData(lastRowCountFormattedData, lastRowCountRawData)

            self.lastFormattedDataRow = lastRowCountFormattedData + 5

        return

    def __getLastRowRawData(self):
        return self.lastRawDataRow

    def __getLastRowFormattedData(self):
        return self.lastFormattedDataRow

    def __updateSheetWithFormattedData(self, currentRowCount, rawDataRowCount):
        # Add Col Headers
        self.formattedDataSheet.write(currentRowCount, 2, "Position Change")
        self.formattedDataSheet.write(currentRowCount, 6, "Net Position")
        self.formattedDataSheet.write(currentRowCount, 9, "Net Change")

        updatedRowCount = currentRowCount + 1

        updatedRowCount = self.__updateSheetWithIndexFuturesData(updatedRowCount, rawDataRowCount)
        updatedRowCount = self.__updateSheetWithIndexCallsData(updatedRowCount, rawDataRowCount)
        updatedRowCount = self.__updateSheetWithIndexPutsData(updatedRowCount, rawDataRowCount)
        updatedRowCount = self.__updateSheetWithStockFuturesData(updatedRowCount, rawDataRowCount)

        return updatedRowCount


    def __updateSheetWithIndexFuturesData(self, currentFormattedRowCount, rawDataRowCount):
        # Add Col Headers
        self.formattedDataSheet.write(currentFormattedRowCount, 1, "Index Futures")
        self.formattedDataSheet.write(currentFormattedRowCount, 2, "Longs")
        self.formattedDataSheet.write(currentFormattedRowCount, 3, "% Change")
        self.formattedDataSheet.write(currentFormattedRowCount, 4, "Shorts")
        self.formattedDataSheet.write(currentFormattedRowCount, 5, "% Change")
        self.formattedDataSheet.write(currentFormattedRowCount, 6, "Today")
        self.formattedDataSheet.write(currentFormattedRowCount, 7, "1 Day Ago")
        self.formattedDataSheet.write(currentFormattedRowCount, 8, "2 Day Ago")
        self.formattedDataSheet.write(currentFormattedRowCount, 9, "Net Change")
        currentFormattedRowCount = currentFormattedRowCount + 1


        # Clients Data
        for index in range (0, 5):
            self.formattedDataSheet.write(currentFormattedRowCount, 1, self.rawDataSheet.cell_value(rawDataRowCount + index - 5, 0))
            self.formattedDataSheet.write(currentFormattedRowCount, 2, float(self.rawDataSheet.cell_value(rawDataRowCount + index - 5, 1)) - float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 1)))
            if float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 1)) != 0:
                self.formattedDataSheet.write(currentFormattedRowCount, 3, (float(self.rawDataSheet.cell_value(rawDataRowCount + index - 5, 1)) - float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 1))) * 100 / float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 1)))
            else:
                self.formattedDataSheet.write(currentFormattedRowCount, 3, 0)
            
            self.formattedDataSheet.write(currentFormattedRowCount, 4, float(self.rawDataSheet.cell_value(rawDataRowCount + index - 5, 2)) - float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 2)))
            if float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 2)) != 0:
                self.formattedDataSheet.write(currentFormattedRowCount, 5, (float(self.rawDataSheet.cell_value(rawDataRowCount + index - 5, 2)) - float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 2)))  * 100 / float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 2)))
            else:
                self.formattedDataSheet.write(currentFormattedRowCount, 5, 0)
            
            self.formattedDataSheet.write(currentFormattedRowCount, 6, float(self.rawDataSheet.cell_value(rawDataRowCount + index - 5, 1)) - float(self.rawDataSheet.cell_value(rawDataRowCount + index - 5, 2)))
            self.formattedDataSheet.write(currentFormattedRowCount, 7, float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 1)) - float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 2)))
            self.formattedDataSheet.write(currentFormattedRowCount, 8, float(self.rawDataSheet.cell_value(rawDataRowCount + index - 19, 1)) - float(self.rawDataSheet.cell_value(rawDataRowCount + index - 19, 2)))
            netChange = (float(self.rawDataSheet.cell_value(rawDataRowCount + index - 5, 1)) - float(self.rawDataSheet.cell_value(rawDataRowCount + index - 5, 2))) - (float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 1)) - float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 2)))
            self.formattedDataSheet.write(currentFormattedRowCount, 9, netChange)
            currentFormattedRowCount = currentFormattedRowCount + 1

        return currentFormattedRowCount


    def __updateSheetWithIndexCallsData(self, currentFormattedRowCount, rawDataRowCount):
        # Add Col Headers
        self.formattedDataSheet.write(currentFormattedRowCount, 1, "Index Calls")
        self.formattedDataSheet.write(currentFormattedRowCount, 2, "Longs")
        self.formattedDataSheet.write(currentFormattedRowCount, 3, "% Change")
        self.formattedDataSheet.write(currentFormattedRowCount, 4, "Shorts")
        self.formattedDataSheet.write(currentFormattedRowCount, 5, "% Change")
        self.formattedDataSheet.write(currentFormattedRowCount, 6, "Today")
        self.formattedDataSheet.write(currentFormattedRowCount, 7, "1 Day Ago")
        self.formattedDataSheet.write(currentFormattedRowCount, 8, "2 Day Ago")
        self.formattedDataSheet.write(currentFormattedRowCount, 9, "Net Change")
        currentFormattedRowCount = currentFormattedRowCount + 1
        
        # Clients Data
        for index in range (0, 5):
            self.formattedDataSheet.write(currentFormattedRowCount, 1, self.rawDataSheet.cell_value(rawDataRowCount + index - 5, 0))
            self.formattedDataSheet.write(currentFormattedRowCount, 2, float(self.rawDataSheet.cell_value(rawDataRowCount + index - 5, 5)) - float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 5)))
            if float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 5)) != 0:
                self.formattedDataSheet.write(currentFormattedRowCount, 3, (float(self.rawDataSheet.cell_value(rawDataRowCount + index - 5, 5)) - float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 5))  ) * 100 / float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 5)))
            else:
                self.formattedDataSheet.write(currentFormattedRowCount, 3, 0)
            self.formattedDataSheet.write(currentFormattedRowCount, 4, float(self.rawDataSheet.cell_value(rawDataRowCount + index - 5, 7)) - float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 7)))
            if float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 7)) != 0:
                self.formattedDataSheet.write(currentFormattedRowCount, 5, (float(self.rawDataSheet.cell_value(rawDataRowCount + index - 5, 7)) - float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 7) ) ) * 100 / float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 7)))
            else:
                self.formattedDataSheet.write(currentFormattedRowCount, 5, 0)
            self.formattedDataSheet.write(currentFormattedRowCount, 6, float(self.rawDataSheet.cell_value(rawDataRowCount + index - 5, 5)) - float(self.rawDataSheet.cell_value(rawDataRowCount + index - 5, 7)))
            self.formattedDataSheet.write(currentFormattedRowCount, 7, float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 5)) - float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 7)))
            self.formattedDataSheet.write(currentFormattedRowCount, 8, float(self.rawDataSheet.cell_value(rawDataRowCount + index - 19, 5)) - float(self.rawDataSheet.cell_value(rawDataRowCount + index - 19, 7)))
            netChange = (float(self.rawDataSheet.cell_value(rawDataRowCount + index - 5, 5)) - float(self.rawDataSheet.cell_value(rawDataRowCount + index - 5, 7))) - (float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 5)) - float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 7)))
            self.formattedDataSheet.write(currentFormattedRowCount, 9, netChange)
            currentFormattedRowCount = currentFormattedRowCount + 1
            
        return currentFormattedRowCount


    def __updateSheetWithIndexPutsData(self, currentFormattedRowCount, rawDataRowCount):
        # Add Col Headers
        self.formattedDataSheet.write(currentFormattedRowCount, 1, "Index Puts")
        self.formattedDataSheet.write(currentFormattedRowCount, 2, "Longs")
        self.formattedDataSheet.write(currentFormattedRowCount, 3, "% Change")
        self.formattedDataSheet.write(currentFormattedRowCount, 4, "Shorts")
        self.formattedDataSheet.write(currentFormattedRowCount, 5, "% Change")
        self.formattedDataSheet.write(currentFormattedRowCount, 6, "Today")
        self.formattedDataSheet.write(currentFormattedRowCount, 7, "1 Day Ago")
        self.formattedDataSheet.write(currentFormattedRowCount, 8, "2 Day Ago")
        self.formattedDataSheet.write(currentFormattedRowCount, 9, "Net Change")
        currentFormattedRowCount = currentFormattedRowCount + 1
        
        # Clients Data
        for index in range (0, 5):
            self.formattedDataSheet.write(currentFormattedRowCount, 1,self.rawDataSheet.cell_value(rawDataRowCount + index - 5, 0))
            self.formattedDataSheet.write(currentFormattedRowCount, 2,float(self.rawDataSheet.cell_value(rawDataRowCount + index - 5, 6)) - float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 6)))
            if float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 6)) != 0:
                self.formattedDataSheet.write(currentFormattedRowCount, 3, (float(self.rawDataSheet.cell_value(rawDataRowCount + index - 5, 6)) - float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 6))  ) * 100 / float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 6)))
            else:
                self.formattedDataSheet.write(currentFormattedRowCount, 3, 0)
            self.formattedDataSheet.write(currentFormattedRowCount, 4,float(self.rawDataSheet.cell_value(rawDataRowCount + index - 5, 8)) - float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 8)))
            if float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 8)) != 0:
                self.formattedDataSheet.write(currentFormattedRowCount, 5, (float( self.rawDataSheet.cell_value(rawDataRowCount + index - 5, 8)) - float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 8))  ) * 100 / float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 8)))
            else:
                self.formattedDataSheet.write(currentFormattedRowCount, 5, 0)
            
            self.formattedDataSheet.write(currentFormattedRowCount, 6,float(self.rawDataSheet.cell_value(rawDataRowCount + index - 5, 6)) - float(self.rawDataSheet.cell_value(rawDataRowCount + index - 5, 8)))
            self.formattedDataSheet.write(currentFormattedRowCount, 7,float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 6)) - float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 8)))
            self.formattedDataSheet.write(currentFormattedRowCount, 8,float(self.rawDataSheet.cell_value(rawDataRowCount + index - 19, 6)) - float(self.rawDataSheet.cell_value(rawDataRowCount + index - 19, 8)))
            netChange = (float(self.rawDataSheet.cell_value(rawDataRowCount + index - 5, 6)) - float(self.rawDataSheet.cell_value(rawDataRowCount + index - 5, 8))) - (float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 6)) - float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 8)))
            self.formattedDataSheet.write(currentFormattedRowCount, 9, netChange)
            currentFormattedRowCount = currentFormattedRowCount + 1
            
        return currentFormattedRowCount


    def __updateSheetWithStockFuturesData(self, currentFormattedRowCount, rawDataRowCount):
        # Add Col Headers
        self.formattedDataSheet.write(currentFormattedRowCount, 1, "Stock Futures")
        self.formattedDataSheet.write(currentFormattedRowCount, 2, "Longs")
        self.formattedDataSheet.write(currentFormattedRowCount, 3, "% Change")
        self.formattedDataSheet.write(currentFormattedRowCount, 4, "Shorts")
        self.formattedDataSheet.write(currentFormattedRowCount, 5, "% Change")
        self.formattedDataSheet.write(currentFormattedRowCount, 6, "Today")
        self.formattedDataSheet.write(currentFormattedRowCount, 7, "1 Day Ago")
        self.formattedDataSheet.write(currentFormattedRowCount, 8, "2 Day Ago")
        self.formattedDataSheet.write(currentFormattedRowCount, 9, "Net Change")
        currentFormattedRowCount = currentFormattedRowCount + 1
        

        # Clients Data
        for index in range (0, 5):
            self.formattedDataSheet.write(currentFormattedRowCount, 1,self.rawDataSheet.cell_value(rawDataRowCount + index - 5, 0))
            self.formattedDataSheet.write(currentFormattedRowCount, 2,float(self.rawDataSheet.cell_value(rawDataRowCount + index - 5, 3)) - float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 3)))
            if float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 3)) != 0:
                self.formattedDataSheet.write(currentFormattedRowCount, 3, (float(self.rawDataSheet.cell_value(rawDataRowCount + index - 5, 3)) - float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 3))  ) * 100 / float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 3)))
            else:
                self.formattedDataSheet.write(currentFormattedRowCount, 3, 0)
            
            self.formattedDataSheet.write(currentFormattedRowCount, 4,float(self.rawDataSheet.cell_value(rawDataRowCount + index - 5, 4)) - float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 4)))
            if float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 4)) != 0:
                self.formattedDataSheet.write(currentFormattedRowCount, 5, (float( self.rawDataSheet.cell_value(rawDataRowCount + index - 5, 4)) - float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 4))  ) * 100 / float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 4)))
            else:
                self.formattedDataSheet.write(currentFormattedRowCount, 5, 0)
            
            self.formattedDataSheet.write(currentFormattedRowCount, 6,float(self.rawDataSheet.cell_value(rawDataRowCount + index - 5, 3)) - float(self.rawDataSheet.cell_value(rawDataRowCount + index - 5, 4)))
            self.formattedDataSheet.write(currentFormattedRowCount, 7,float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 3)) - float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 4)))
            self.formattedDataSheet.write(currentFormattedRowCount, 8,float(self.rawDataSheet.cell_value(rawDataRowCount + index - 19, 3)) - float(self.rawDataSheet.cell_value(rawDataRowCount + index - 19, 4)))
            netChange = (float(self.rawDataSheet.cell_value(rawDataRowCount + index - 5, 3)) - float(self.rawDataSheet.cell_value(rawDataRowCount + index - 5, 4))) - (float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 3)) - float(self.rawDataSheet.cell_value(rawDataRowCount + index - 12, 4)))
            self.formattedDataSheet.write(currentFormattedRowCount, 9, netChange)
            currentFormattedRowCount = currentFormattedRowCount + 1
            
        return currentFormattedRowCount

    def __updateCellFloat(self, sheet, row, col, data):
        try:
            sheet.write(row, col, float(data))
        except:
            sheet.write(row, col, 0)
        return 