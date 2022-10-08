#STOCK DATA-90 Days from now
#---------BY------------
#SANOJ KUMAR PRADHAN
#-----------------------

import yfinance as yf  #imported yfinance library to extract data from yahoo finance API
import pandas as pd  #imported pandas library to create dataframes
from datetime import date, timedelta  #imported this library to set start and end date.

excel_data=pd.read_excel('Symbols.xlsx')   #Read the excel file containing Stock Symbols to pandas dataframe.
col_data=list(excel_data["Symbol"])        #retrieved the values of column_name-"Symbol" and stored in a list.

#Time settings
start_date = (date.today()-timedelta(days=90)).isoformat()
end_date = date.today().isoformat()


#iterated through the list of column data containing Symbols
for i in col_data:
    try:
        if i=="PCHFL":
            continue

        ticker=f'{i}.NS'
        print(ticker)
        data = yf.download(ticker, start_date, end_date)    #retrieved data from the yahoo finance API
        data["Date"]=data.index                             #set "Date" column as the index
        data.index = data.index.strftime('%Y-%m-%d %H:%M:%S') #string formatted the index values
        data = data[["Close"]]
        data.to_excel(f'Stock_info_in_excel/{i}.xlsx', sheet_name="sheet 1",index=True)  #stored the data obtained in excel format
    except AttributeError:
        continue


#DEBUGGING_CHANGES

#CADILA HEALTHCARE LIMITED           ,ZYDUSLIFE(symbol updated)
#HEXAWARE TECHNOLOGIES LTD           ,HEXAWARE(has been delisted)- handled the AttributeError using "try-except"
#DEWAN HOUSING FIN CORP LTD          ,PCHFL(No data found) - Bypassed the Symbol using "if-statement"
#MOTHERSON SUMI SYSTEMS LT           ,MOTHERSON(symbol updated)
#BHARTI INFRATEL LTD.                ,INDUSTOWER(symbol updated)
#TATA GLOBAL BEVERAGES LTD           ,TATACONSUM(symbol updated)
#NIIT TECHNOLOGIES LTD               ,COFORGE(symbol updated)


