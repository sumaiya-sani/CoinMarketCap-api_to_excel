import json
import requests
import xlsxwriter


local_curency="USD"
local_symbol="$"

api_key=''
headers={'Accepts': 'application/json',
'X-CMC_PRO_API_KEY':api_key}
base_url='https://pro-api.coinmarketcap.com'


crypto_workbook=xlsxwriter.Workbook('cryptocurrencies.xlsx') #crypto_workbook > to write new xslx project ,#xlsxwriter.workbook(اساسي)
crypto_sheet=crypto_workbook.add_worksheet() #cryptosheet> to define a new sheet # add_worksheet (اساسي من المكتبة )

crypto_sheet.write('A1','Name')
crypto_sheet.write('B1','Symbol')
crypto_sheet.write('C1','Markrt Cap')
crypto_sheet.write('D1','Price')
crypto_sheet.write('E1','24H Volume')
crypto_sheet.write('F1','Hour Change')
crypto_sheet.write('G1','Day Change')
crypto_sheet.write('H1','Week Change')
#here we going to start a loop but every time he loop he only show us a 1000 crypto info 
start= 1
row = 1 
for i in range(10):
#start in url which number we wanna start with 
    listing_url=base_url+'/v1/cryptocurrency/listings/latest?convert='+local_curency+'&start='+str(start)
    request=requests.get(listing_url,headers=headers)
    results=request.json()
    data=results['data']
#we loops in the data that's we get from the results 
#no need to dfine data as a variable
    for currency in data :
        name=currency['name']
        symbol=currency['symbol']

        quote=currency['quote'][local_curency]
        market_cap=quote['market_cap']
        hour_change=quote['percent_change_1h']
        day_change=quote['percent_change_24h']
        week_change=quote['percent_change_7d']
        price=quote['price']
        volume=quote['volume_24h']

        volume_string='{:,}'.format(volume)
        market_cap_string='{:,}'.format(market_cap)

        #to add it to the row in the cheet 
        crypto_sheet.write(row,0,name)
        crypto_sheet.write(row,1,symbol)
        crypto_sheet.write(row,2,local_symbol+market_cap_string)
        crypto_sheet.write(row,3,local_symbol+str(price))
        crypto_sheet.write(row,4,local_symbol+volume_string)
        crypto_sheet.write(row,5,str(hour_change)+"%")
        crypto_sheet.write(row,6,str(day_change)+"%")
        crypto_sheet.write(row,7,str(week_change)+"%")

        row +=1 
        #print(json.dumps(data,sort_keys=True,indent=4))
        start+=100
crypto_workbook.close() #to save all what we write in the sheet     


