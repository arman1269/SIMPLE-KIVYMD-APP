from ast import Global
from ctypes.wintypes import DWORD
from lib2to3.pytree import convert
from urllib import response
from urllib.request import urlopen
from kivymd.app import MDApp
from kivy.uix.screenmanager import Screen,ScreenManager
from kivy.lang import Builder
from kivymd.uix.list import MDList,ThreeLineListItem
from kivymd.uix.dialog import MDDialog
from kivymd.uix.button import MDFlatButton, MDRectangleFlatButton
from kivymd.uix.datatables import MDDataTable
from kivy.metrics import dp
from kivy.uix.recycleview import RecycleView
from kivy.uix.scrollview import ScrollView
from kivy.core.window import Window
from kivy.uix.image import Image
import tkinter as tk
import numpy as np 

import pygsheets
import pandas as pd

import requests
import json
import locale
from datetime import datetime
from plyer import notification

locale.setlocale(locale.LC_ALL, '')

class Spk(Screen):
    pass        
        # initial = 0
        
        # def on_touch_down(self, touch):
        #     self.initial = touch.x
        #     return super(Screen, self).on_touch_down(touch) #pass the touch on 

        # def on_touch_up(self, touch):

        #     if touch.x > self.initial:
                
        #         self.addbar()

        #         return True #don't pass the touch on
        
        #     elif touch.x < self.initial:

        #         self.addbar()

        #         return True #don't pass the touch on
        #     else: 

        #         return super(Screen, self).on_touch_up(touch)

        # def addbar(self):

        #     global i
        #     spklistadd = ThreeLineListItem(text = response[i]['PRNo']+'/'+response[i]['Code'],secondary_text= response[i]['CreateDate'][0:10], tertiary_text = response[i]['Supplier'], on_release=self.print_item2)

        #     self.ids.ls.add_widget(spklistadd)
                
        #     i+=1        
class List(Screen):
    pass
                
class ScreenManager(ScreenManager):
    pass

class Detail(Screen):
    def back(self):
        self.manager.current='list'
    def approved(self):
        spkApp.approve(self)

x = ''


class spkApp(MDApp):
    dialog = None
    def build(self):
        
        screen = Screen()

        self.kave = Builder.load_file('main.kv')  
        
        screen.add_widget(self.kave)
        return screen
    
    # def on_start(self):
        
    #     client = pygsheets.authorize(service_file='creds.json')

    #     spreadsheet_id = "15KtAVhAwn1gpfMv6wZTdXKQxnoXKOwyGUmizuMcQBhk" # Please set your Spreadsheet ID.

    #     sh = client.open('SPK')

    #     # #select the first sheet 
    #     wks = sh[1]

    #     sh = client.open_by_key(spreadsheet_id)

    #     df = wks.get_as_df(include_tailing_empty_rows=False)

    #     filtered_df = df.loc[df['CodeStatus'] == 4]
 
    #     # print(df['Code'].where(df['CodeStatus'] == 4))
        

    #     def on_tampil():
    #         for index,row in filtered_df.iterrows():
    #             global spklist
    #             spklist = ThreeLineListItem(text = str(row[0]) +'/'+ str(row[1])[0:10],secondary_text= row[2], tertiary_text = row[3], on_release=self.print_item)
    #             self.kave.get_screen('spk').ids.ls.add_widget(spklist)
          
    #     on_tampil()     

        
    def print_item(self, instance):
        self.kave.get_screen('detail').manager.current = 'detail'  
        x = instance.text.split('/')        
        global code
        # code = {'Code' : x[0]}
        code = x[0]

        client = pygsheets.authorize(service_file='creds.json')

        spreadsheet_id = "15KtAVhAwn1gpfMv6wZTdXKQxnoXKOwyGUmizuMcQBhk" # Please set your Spreadsheet ID.

        sh = client.open('SPK')

        # #select the first sheet 
        wks = sh[2]

        sh = client.open_by_key(spreadsheet_id)

        # df = wks.get_as_df(include_tailing_empty_rows=False)
        data_row = []

        cells = wks.find(code, searchByRegex=False, matchCase=True, 
            matchEntireCell=True, includeFormulas=False, 
            cols=(1,1), rows=None, forceFetch=True)
        
        for cell in cells:
            data_row.append(wks.get_row(cell.row, returnas='matrix', include_tailing_empty=False))
        
        data_row.append([" ", " ", " ", " ", " ", " "])   
        table = None
        
        table = MDDataTable(
            pos_hint = {'center_x': 0.5, },
            size_hint = (0.9, 0.7),
            # use_pagination = True,
            # rows_num=10,
            column_data = [
                ("SPK#", dp(10)),
                ("Description", dp(60)), 
                ("Qty", dp(30)),
                ("Unit", dp(30)),
                ("Price", dp(30)),
                ("Total", dp(30))
            ],

            row_data = data_row
            
        )

        # print (data_row)
        self.kave.get_screen('detail').ids.blok1.clear_widgets()
        self.kave.get_screen('detail').ids.blok1.add_widget(table)

    def approve(self):
        client = pygsheets.authorize(service_file='creds.json')

        spreadsheet_id = "15KtAVhAwn1gpfMv6wZTdXKQxnoXKOwyGUmizuMcQBhk" # Please set your Spreadsheet ID.

        sh = client.open('SPK')

        # #select the first sheet 
        sh = client.open_by_key(spreadsheet_id)
        wks = sh[3]

        df = pd.DataFrame()

        CodeSPK = []
        ApprovedDate = []
        ApprovedBy = []

        today = datetime.now()

        CodeSPK.append(code)
        ApprovedBy.append('diyu')
        ApprovedDate.append(today)

        values = wks.get_values(start='A1', end='A500000')
        num_rows = len(values)

        df['CodeSPK'] = CodeSPK
        df['Date'] = ApprovedDate
        df['ApprovedBy'] = ApprovedBy

        wks.set_dataframe(df,(1,1))

        if num_rows == 1:
            wks.set_dataframe(df, (0,1), copy_index=False, copy_head=True, extend=True, fit=False, escape_formulae=False)
        else:
            wks.set_dataframe(df, (num_rows+1,1), copy_index=False, copy_head=False, extend=True, fit=False, escape_formulae=False)

        update_spk()  
        self.manager.current='spk'

    def search(self):
        # self.kave.get_screen('spk').ids.ls.clear_widgets()
        codespk= self.kave.get_screen('spk').ids.sr.text
        self.kave.get_screen('spk').ids.ls.clear_widgets()

        client = pygsheets.authorize(service_file='creds.json')

        spreadsheet_id = "15KtAVhAwn1gpfMv6wZTdXKQxnoXKOwyGUmizuMcQBhk" # Please set your Spreadsheet ID.

        sh = client.open('SPK')
        
        # #select the first sheet 
        wks = sh[1]

        global spklist
        df = wks.get_as_df(include_tailing_empty_rows=False)

        fdf = df.loc[(df['Code'] <= int(codespk)) & (df['CodeStatus'] == 4)]

        for index,row in fdf.iterrows():
            spklist = ThreeLineListItem(text = str(row[0]) +'/'+ str(row[1])[0:10],secondary_text= row[2], tertiary_text = row[3], on_release=self.print_item)
            self.kave.get_screen('spk').ids.ls.add_widget(spklist)
    
    def login(self,users, passws):

        client = pygsheets.authorize(service_file='creds.json')

        spreadsheet_id = "15KtAVhAwn1gpfMv6wZTdXKQxnoXKOwyGUmizuMcQBhk" # Please set your Spreadsheet ID.

        sh = client.open('SPK')

        # #select the first sheet 
        wks = sh[4]

        sh = client.open_by_key(spreadsheet_id)

        df = wks.get_as_df(include_tailing_empty_rows=False)

        # filtered_df = df.loc[(df['UserName'] == 'arma') & (df['Password'] == 2)]
        filtered_df = df.loc[:, users:passws]

        if filtered_df.empty:
            # print('user not found!')
            notification.notify(title='SPK', message='User not found!')
        else:
            self.kave.get_screen('detail').manager.current = 'list'  
            
            client = pygsheets.authorize(service_file='creds.json')

            spreadsheet_id = "15KtAVhAwn1gpfMv6wZTdXKQxnoXKOwyGUmizuMcQBhk" # Please set your Spreadsheet ID.

            sh = client.open('SPK')

            # #select the first sheet 
            wks = sh[1]

            sh = client.open_by_key(spreadsheet_id)

            df = wks.get_as_df(include_tailing_empty_rows=False)

            filtered_df = df.loc[df['CodeStatus'] == 4]
 
            for index,row in filtered_df.iterrows():
                global spklist
                spklist = ThreeLineListItem(text = str(row[0]) +'/'+ str(row[1])[0:10],secondary_text= row[2], tertiary_text = row[3], on_release=self.print_item)
                self.kave.get_screen('list').ids.ls.add_widget(spklist)
            

def update_spk():
    client = pygsheets.authorize(service_file='creds.json')

    spreadsheet_id = "15KtAVhAwn1gpfMv6wZTdXKQxnoXKOwyGUmizuMcQBhk" # Please set your Spreadsheet ID.

    sh = client.open('SPK')

    # #select the first sheet 
    sh = client.open_by_key(spreadsheet_id)
    wks = sh[1]
    nospk = wks.find(code)    
    cellupdate = 'E'+str(nospk[0].row)
    wks.cell(cellupdate).value = "0"
              
        
        
spkApp().run()