from datetime import date
from itertools import count
import requests
from GlobalConstants import GlobalConstants
from NseDataModel import NseHistoricalData, ParticipantWiseNseRawRecords
from nsepy import get_history


class NseOptionDataRepository:

    def __init__(self):
        self.globalConstants = GlobalConstants()

    def getParticipantData(self, date) -> ParticipantWiseNseRawRecords:
        output = ParticipantWiseNseRawRecords([], [], [], [], [], [], [])

        # fetch data from nse
        url = self.globalConstants.getNseParticipantDataUrl(date)
        r = requests.get(url)
        outputData = []
        if r.status_code != 200:
            return None, False
        else:
            tempOutput = r.content
            outputData = tempOutput.splitlines()

        dateRow = True
        headerRow = True
        count = 0
        for rowDataStr in outputData:
            if dateRow:
                rowData = ['', 'Participant wise Open Interest',
                           '', '', 'Date - ' + date]
                output.dateHeader = rowData
                dateRow = False
            elif headerRow:
                rowDataTemp = str(rowDataStr).split(',')
                rowData = [s.replace("b'", "") for s in rowDataTemp]
                output.columnHeader = rowData
                headerRow = False
            else:
                rowDataTemp = str(rowDataStr).split(',')
                heading = rowDataTemp[0].replace("b'", "")
                rowDataTemp.remove(rowDataTemp[0])
                rowData = [heading]
                for rawCellData in rowDataTemp:
                    rowData.append(str(rawCellData.replace("'", "")))
                # format_cell_range(sheet, "A"+str(startIndex)+":Z"+str(startIndex), fmtData)
                if count == 0:
                    output.clientData = rowData
                    count = count + 1
                elif count == 1:
                    output.diiData = rowData
                    count = count + 1
                elif count == 2:
                    output.fiiData = rowData
                    count = count + 1
                elif count == 3:
                    output.proData = rowData
                    count = count + 1
                elif count == 4:
                    output.totalData = rowData

        return output, True

    def getNiftyHistoricalData(self, inputDate: str):
        dateData = inputDate.split('-')
        x = int(dateData[2])
        x = int(dateData[1])
        x = int(dateData[0])
        tempDate = date(int(dateData[2]), int(dateData[1]), int(dateData[0]))
        data = get_history(symbol="NIFTY", start=tempDate,
                           end=tempDate, index=True).values
        if len(data) == 0:
            return None, False
        else:
            return NseHistoricalData(
                data[0][0], data[0][3], data[0][1], data[0][2]), True
