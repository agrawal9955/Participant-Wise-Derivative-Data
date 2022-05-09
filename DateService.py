from datetime import date, timedelta
import datetime
from time import time
from GlobalConstants import GlobalConstants
from GoogleSheetService import GoogleSheetService


class DateService:
    def __init__(self, sheetService: GoogleSheetService) -> None:
        self.googleSheetService = sheetService
        self.globalConstants = GlobalConstants()
        return

    def fetchApplicableDate(self):
        lastDataDate = self.googleSheetService.getLastDate().split("-")
        currentDate = self.__getCurrentDate().split("-")
        if self.__checkCurrentDateWrtLastDate(currentDate, lastDataDate):
            d = date(int(lastDataDate[2]), int(lastDataDate[1]), int(
                lastDataDate[0])) + timedelta(days=1)
            return d.strftime("%d-%m-%Y"), True
        else:
            return "", False

    def __getCurrentDate(self) -> str:
        # data will be updated after 9PM
        now = datetime.datetime.now()
        currentTime = now.strftime("%H-%M").split("-")
        marketDataTime = self.globalConstants.getMarketDataAvailableTime().split("-")
        if int(currentTime[0]) >= int(marketDataTime[0]) and int(currentTime[1]) >= int(marketDataTime[1]):
            today = date.today()
        else:
            today = date.today() - timedelta(days=1)

        d1 = today.strftime("%d-%m-%Y")
        return d1

    def __checkCurrentDateWrtLastDate(self, current, last) -> bool:
        if int(current[2]) <= int(last[2]):
            if int(current[1]) <= int(last[1]):
                if int(current[0]) <= int(last[0]):
                    return False

        return True
