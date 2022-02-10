#подключаем библиотеки

import shelve
import pandas as pd
import requests
from bs4 import BeautifulSoup as bs

#ссылка на страницу заказов
ORDERS_LINK = 'https://trade.aliexpress.com/orderList.htm'

#открываем файл данных с токенами
tokensExel = pd.read_excel('Tokens.xlsx', header=None)

dataList = []   #список данных, которые парсяться на странице аккаунта
tokenList = []  #список токенов

#считываем токены с таблицы
for i in tokensExel.values.tolist():
    tokenList.append(i[0])

#проходимся по каждому токену
for token in tokenList:
    print('Аккаунт ' + str(tokenList.index(token)+1) + '/' + str(len(tokenList)+1))
    try:
        r = None
        cookie = {'session': '17ab96bd8ffbe8ca58a7865'}

        #формируем сессию
        with requests.Session() as session:
            session.post(token, cookie)
            r = session.get(ORDERS_LINK)
            
            #начинаем порсить страницу
            soup = bs(r.text, "lxml")  
            
            #ищем все контейнеры с заказами
            ordersValue = soup.find_all('tbody', class_='order-item-wraper')
            i = 0
            
            #проходимся по заказам
            for order in ordersValue:
                orderStatus = order.find('td', class_='order-status').find('span', class_='f-left').text.strip()
                orderName = order.find('a',  class_= 'baobei-name').text.strip()
                orderPrice = order.find('p',  class_= 'amount-num').text.strip()

                trackCode = ''

                #если статус "доставляеться", пытаемся получить трек-номер
                if orderStatus == 'Awaiting delivery':
                    trackCodeLink = 'https://trade.aliexpress.com/' + order.find('a', class_='view-detail-link').get('href')
                    
                    r2 = session.get(trackCodeLink)
                    soup2 = bs(r2.text, "lxml")  
                    trackCode = soup2.find('td', class_='no').find('div').text.strip()

                #добавляем данные в базу
                dataList.append([token, orderStatus, trackCode, i, orderName, orderPrice])
                i = i+1
    except:
        print('Токен не валид')

#список данных с базы
lastDataList = []

#пытаемся открыть базу данных
try:
	shalveData = shelve.open('database')
	lastDataList = shalveData['data']
	shalveData.close()
except:
    print('Видимо первый запуск')

#формируем список измененных статусов
changedData = []
for lastData in lastDataList:
    for curData in dataList:
        if lastData[0] == curData[0] and lastData[3] == curData[3]:     #если совпадает токен и  идентификатор
            if lastData[1] != curData[1]:       #если статусы заказов отличаються
                changedData.append([lastData[0], lastData[1], curData[1], curData[2], curData[4], curData[5]])

#формируем таблицу всех данных
dfToken = []
dfValue = []
dfTrack = []
dfName = []
dfPrice = []

for curData in dataList:        
    dfToken.append(curData[0])
    dfValue.append(curData[1])
    dfTrack.append(curData[2])
    dfName.append(curData[4])
    dfPrice.append(curData[5])

df = pd.DataFrame({'Токен': dfToken,
                   'Статус': dfValue,
                   'Трек': dfTrack,
                   'Имя': dfName,
                   'Цена': dfPrice})

df.to_excel('AllData.xlsx', sheet_name='Данные', index=False)

#формируем таблицу измененных данных
dfToken = []
dfOldValue = []
dfNewValue = []
dfTrack = []
dfName = []
dfPrice = []

for change in changedData:        
    dfToken.append(change[0])
    dfOldValue.append(change[1])
    dfNewValue.append(change[2])
    dfTrack.append(change[3])
    dfName.append(curData[4])
    dfPrice.append(curData[5])

df = pd.DataFrame({'Токен': dfToken,
                   'Старое': dfOldValue,
                   'Обнова': dfNewValue,
                   'Трек': dfTrack,
                   'Имя': dfName,
                   'Цена': dfPrice})

df.to_excel('Changed.xlsx', sheet_name='Данные', index=False)

#перезаписываем базу данных
shalveData = shelve.open('database')
shalveData['data'] = dataList
shalveData.close()