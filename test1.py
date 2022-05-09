from datetime import date
import pandas as pd
from nsepy import get_history

netChangeIndex = 9
dataStartIndex = 0
dateIndex = 0

df = pd.read_excel('Market_Stats_Python_Managed.xlsx',
                   sheet_name='Flat Data')

data = df.values[0:]

dates = []
openPrices = []
closePrices = []
highPrices = []
lowPrices = []
excelMap = {}

for index in range(dataStartIndex, len(data)):
    dateData = str(data[index][0]).split("-")
    candleData = get_history(symbol="NIFTY", start=date(
        int(dateData[2]), int(dateData[1]), int(dateData[0])), end=date(
        int(dateData[2]), int(dateData[1]), int(dateData[0])), index=True)
    dates.append(str(data[index][0]))
    candleData = candleData.values[0]
    openPrices.append(float(candleData[0]))
    closePrices.append(float(candleData[3]))
    highPrices.append(float(candleData[1]))
    lowPrices.append(float(candleData[2]))
    print(str(data[index][0]))

excelMap["NIFTY OPEN"] = openPrices
excelMap["NIFTY CLOSE"] = closePrices
excelMap["NIFTY HIGH"] = highPrices
excelMap["NIFTY LOW"] = lowPrices
excelMap["Date"] = dates


outputDF = pd.DataFrame(excelMap)
outputDF.to_excel("index_data.xlsx")
