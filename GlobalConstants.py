import enum


class GlobalConstants():

    def __init__(self) -> None:
        self.NSE_URL = "https://www1.nseindia.com/content/nsccl/fao_participant_oi_REQUEST__DATE.csv"
        self.SHEET_NAME = "Market_Stats_Python_Managed"

    def getNseParticipantDataUrl(self, date):
        return self.NSE_URL.replace("REQUEST__DATE", date.replace('-', ''))

    def getGoogleWorksheetName(self):
        return self.SHEET_NAME

    def getMarketDataAvailableTime(self):
        # 9PM IST
        return "21-0"

    def getFormattedDataSheet(self):
        # Sheet No 2
        return 3

    def getRawDataSheet(self):
        # Sheet No 3
        return 1

    def getMetaDataSheet(self):
        # Sheet No 4
        return 0

    def getFormattedFlatDataSheet(self):
        # Sheet No 5
        return 4
