from datetime import datetime
from time import sleep
from DateService import DateService
from ExcelSheetService import ExcelSheetService
from GoogleSheetService import GoogleSheetService
from NseDataRepo import NseOptionDataRepository


nseRepo = NseOptionDataRepository()
googleSheetService = GoogleSheetService()
excelService = ExcelSheetService()
dateService = DateService(googleSheetService)

while True:
    date, status = dateService.fetchApplicableDate()
    if not status:
        print(datetime.now().strftime("%H-%M") + " - sleeping for 1 hour")
        sleep(3600)
        continue
    data, status = nseRepo.getParticipantData(date)
    if not status:
        googleSheetService.setLastDate(date)
        continue
    niftyData, status = nseRepo.getNiftyHistoricalData(date)
    if not status:
        continue
    googleSheetService.addParticipantRawData(data)
    googleSheetService.addParticipantDataFormula()
    googleSheetService.addParticipantFormattedFlatData(date, niftyData)
    googleSheetService.setLastDate(date)
    print("data added for date: " + date)
