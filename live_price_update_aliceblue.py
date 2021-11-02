import pandas as pd
import xlwings as xw
from alice_blue import *
import datetime
import sys, traceback,os
import win32com.client as client
import logging        
import pandas as pd
import pandas_datareader.data as web
import numpy as np
import datetime as dt
from datetime import date,datetime,timedelta
import requests
import time


api_secret=''
access_token = AliceBlue.login_and_get_access_token(username= "",
                                                    password= "",
                                                    twoFA='',
                                                    api_secret=api_secret,
                                                    app_id='')

print("Access Token generated is :", access_token)

alice = AliceBlue(username=user_name,
                  password=pass_word,
                  access_token=access_token,
                  master_contracts_to_download=['NSE', 'BSE','MCX','NFO'])


def subscribe_symbol(name):
    if exchange != 'NFO':
        instrument_name = alice.get_instrument_by_symbol('NSE', symbol = name)
    instrument=alice.subscribe(instrument_name, LiveFeedType.MARKET_DATA)
        
    
def process_xlsx():    
    for x in range(3,max_row):
        symbol_pos = symbol_position +str(x)
        ltp_pos = ltp_position + str(x)
        name_pos = name_position + str(x)
        active_rows_val = list(active_rows.keys())
        if x not in active_rows_val:            
            if sht.range(name_pos).value != None:
                name = sht.range(name_pos).value
                symbol = sht.range(symbol_pos).value
                active_rows[x] = name
                if symbol not in list(xlsx_mapping.keys()):
                    name_name = sht.range(name_pos).value
                    if name_name != None:
                        xlsx_mapping[symbol] = [ltp_pos]
##                        print('name_{}'.format(name_name))
                        subscribe_symbol(name_name)
                        
                elif ltp_pos not in xlsx_mapping[symbol]:
                    xlsx_mapping[symbol].append(ltp_pos)
        else:
            if sht.range(symbol_pos).value == None:
                symbol = sht.range(symbol_pos).value
                unsubscribe(x)    
        
def is_symbol_in_xl_mapping(symbol):
    
    if symbol not in list(xlsx_mapping.keys()):
        return False
    return True

def update_pos(arguments):
        ltp = arguments['data']
        pos = arguments['location']
        sht.range(pos).value = ltp

def unsubscribe(row_num):
    symbol = active_rows[row_num]
    if len(xlsx_mapping[symbol]) == 1:
        instrument_name = alice.get_instrument_by_symbol(exchange, symbol = symbol)
        alice.unsubscribe(instrument_name, LiveFeedType.MARKET_DATA)
        del xlsx_mapping[symbol]
    else:
        for pos in list(xlsx_mapping[symbol]):
            if str(pos[1]) == str(row_num):
                sht.range(pos).value = None
                xlsx_mapping[symbol].remove(pos)
    del active_rows[row_num]

def event_handler_quote_update(message):
    symbol = message['instrument'].symbol
    symbol_ltp_xlsx_position_list = xlsx_mapping[symbol]
    for position in symbol_ltp_xlsx_position_list:
        for data in list(all_data):
            if position == data['location']:
                all_data.remove(data)
        all_data.append({'location':position, 'data':message['ltp']})

def open_callback():
    global socket_opened
    socket_opened = True
    
if __name__ == '__main__':

    max_row = 33
    exchange = 'NSE'
    xlsx_mapping = {}
    socket_opened = False
    wb=xw.Book('intraday_stocks.xlsx')
    sht=wb.sheets['Sheet1']
    name_position = 'A'
    symbol_position = 'A'
    ltp_position = 'B'
    active_rows = {}
    all_data = []
    
    start=dt.datetime(2021,4,8)
    sht.range('D'+str(1)).value='Current time'
    alice.start_websocket(subscribe_callback=event_handler_quote_update,
                        socket_open_callback=open_callback,
                        run_in_background=True)

    while(socket_opened==False):
        pass

    process_xlsx()

    while True:

        ## handle event update (new value reveieved from websocket)
        for i in range(len(all_data)):
            try:
                update_pos(all_data[i])
            except Exception as e:
                print(e)
        
        ## check xl for any new symbols or deletion of old symbol.
        process_xlsx()
