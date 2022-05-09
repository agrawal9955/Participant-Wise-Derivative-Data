import pandas as pd

netChangeIndex = 9
dataStartIndex = 1

df1 = pd.read_excel('Market_Stats_Python_Managed.xlsx',
                    sheet_name='Daily OI Analysis - Formatted D')
df2 = pd.read_excel('Market_Stats_Python_Managed.xlsx',
                    sheet_name='Sheet10')

data1 = df1.values[0:]
data2 = df2.values[0:]

IndexFuturesClientData = []
IndexFuturesDiiData = []
IndexFuturesFiiData = []
IndexFuturesProData = []
IndexCallsClientData = []
IndexCallsDiiData = []
IndexCallsFiiData = []
IndexCallsProData = []
IndexPutsClientData = []
IndexPutsDiiData = []
IndexPutsFiiData = []
IndexPutsProData = []
StockFuturesClientData = []
StockFuturesDiiData = []
StockFuturesFiiData = []
StockFuturesProData = []
excelMap = {}

for index in range(dataStartIndex, len(data1), 30):
    IndexFuturesClientData.append(data1[index][netChangeIndex])
    IndexFuturesDiiData.append(data1[index+1][netChangeIndex])
    IndexFuturesFiiData.append(data1[index+2][netChangeIndex])
    IndexFuturesProData.append(data1[index+3][netChangeIndex])

    IndexCallsClientData.append(data1[index+6][netChangeIndex])
    IndexCallsDiiData.append(data1[index+7][netChangeIndex])
    IndexCallsFiiData.append(data1[index+8][netChangeIndex])
    IndexCallsProData.append(data1[index+9][netChangeIndex])

    IndexPutsClientData.append(data1[index+12][netChangeIndex])
    IndexPutsDiiData.append(data1[index+13][netChangeIndex])
    IndexPutsFiiData.append(data1[index+14][netChangeIndex])
    IndexPutsProData.append(data1[index+15][netChangeIndex])

    StockFuturesClientData.append(data1[index+18][netChangeIndex])
    StockFuturesDiiData.append(data1[index+19][netChangeIndex])
    StockFuturesFiiData.append(data1[index+20][netChangeIndex])
    StockFuturesProData.append(data1[index+21][netChangeIndex])
for index in range(dataStartIndex, len(data2), 30):
    IndexFuturesClientData.append(data1[index][netChangeIndex])
    IndexFuturesDiiData.append(data1[index+1][netChangeIndex])
    IndexFuturesFiiData.append(data1[index+2][netChangeIndex])
    IndexFuturesProData.append(data1[index+3][netChangeIndex])

    IndexCallsClientData.append(data1[index+6][netChangeIndex])
    IndexCallsDiiData.append(data1[index+7][netChangeIndex])
    IndexCallsFiiData.append(data1[index+8][netChangeIndex])
    IndexCallsProData.append(data1[index+9][netChangeIndex])

    IndexPutsClientData.append(data1[index+12][netChangeIndex])
    IndexPutsDiiData.append(data1[index+13][netChangeIndex])
    IndexPutsFiiData.append(data1[index+14][netChangeIndex])
    IndexPutsProData.append(data1[index+15][netChangeIndex])

    StockFuturesClientData.append(data1[index+18][netChangeIndex])
    StockFuturesDiiData.append(data1[index+19][netChangeIndex])
    StockFuturesFiiData.append(data1[index+20][netChangeIndex])
    StockFuturesProData.append(data1[index+21][netChangeIndex])

excelMap["Index Futures Client"] = IndexFuturesClientData
excelMap["Index Futures DII"] = IndexFuturesDiiData
excelMap["Index Futures FII"] = IndexFuturesFiiData
excelMap["Index Futures Pro"] = IndexFuturesProData

excelMap["Index Calls Client"] = IndexCallsClientData
excelMap["Index Calls DII"] = IndexCallsDiiData
excelMap["Index Calls FII"] = IndexCallsFiiData
excelMap["Index Calls Pro"] = IndexCallsProData

excelMap["Index Puts Client"] = IndexPutsClientData
excelMap["Index Puts DII"] = IndexPutsDiiData
excelMap["Index Puts FII"] = IndexPutsFiiData
excelMap["Index Puts Pro"] = IndexPutsProData

excelMap["Stock Futures Client"] = StockFuturesClientData
excelMap["Stock Futures DII"] = StockFuturesDiiData
excelMap["Stock Futures FII"] = StockFuturesFiiData
excelMap["Stock Futures Pro"] = StockFuturesProData

outputDF = pd.DataFrame(excelMap)
outputDF.to_excel("market_data_flat.xlsx")
