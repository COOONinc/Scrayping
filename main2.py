from logging import error
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from pyasn1.type.univ import Null 
from selenium import webdriver
from time import sleep
import datetime
from selenium.webdriver.common.keys import Keys
#ActionChainsを使う時は、下記のようにActionChainsのクラスをロードする
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from bs4 import BeautifulSoup, BeautifulStoneSoup
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import requests
import json

# mac環境用
import chromedriver_binary

class Scrayping:
    #初期化処理
    def __init__(self):
        self.options = Options()
        #ブラウザ立ち上げは処理が重たくなるので本番環境ではheadlessモードを採用
        self.options.add_argument('--headless')
        self.options.add_argument('--window-size=1920,1080')
        #mac環境用
        self.driver = webdriver.Chrome(options=self.options)

    def act(self,url):
        self.driver.get(url)

        # 企業リスト全件取得
        tableElem =self.driver.find_element_by_xpath('//*[@id="01"]')
        trs = tableElem.find_elements(By.TAG_NAME, "li")

        # 企業情報格納リスト
        Name = []
        Website = []
        Phone = []

        start = 171 #change

        for num in range(start,200): #change
            self.driver.get(url)
            print(num)
            try:
              link = self.driver.find_element_by_xpath(f'//*[@id="01"]/ul/li[{num}]/a')
              print(link)
              # 各企業のURL遷移
              link.click()
              # 別タブ遷移型なのでwindowを変更
              handle_array = self.driver.window_handles # 遷移したタブを配列に格納
              print(f'{num}番目')
              print(handle_array)
              print(handle_array[num-start-1])
              self.driver.switch_to.window(handle_array[num-start-1]) #末尾に追加されていくのでindex[-1]の値を取得
              sleep(5)

              page_source = self.driver.page_source
              soup = BeautifulSoup(page_source, 'html.parser')

              inc = soup.select_one('div.ttltext a')
              info = soup.select_one('div.search_detail_info tr:nth-of-type(2) td:nth-of-type(2)')

              if(inc is not None):
                name = inc.string
                website = inc.get('href')
              if(info is not None):
                phone = info.text
              else:
                phone = None

              

              # 値が入っていない企業には空文字をセットする
              if(name is not None):
                  Name.append(name)
              else:
                  Name.append('')
              if(website is not None):
                  Website.append(website)
              else:
                  Website.append('')
              if(phone is not None):
                Phone.append(phone)
              else:
                Phone.append('')
              
              print(Name)
              print(Website)
              print(Phone)
  
              print(f'現在:{num}社目')
            except:
              print(f'linkが見つかりません。スキップします')


        # --- ここからスクレイピング結果を出力するコード ---

        # # Use a service account
        scope = ['https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive']
  
        # #jsonファイルを指定
        credentials = ServiceAccountCredentials.from_json_keyfile_name('serviceAccount.json', scope)
  
        # # 認証
        gc = gspread.authorize(credentials)

        # 「キー」でワークブックを取得
        # https://docs.google.com/spreadsheets/d/{{SPREADSHEET_KEY}}/edit?usp=sharing
        # ※必ず、共有からリンクを知っている人は閲覧可能に変更しておくこと
        # ※GCPのサービスアカウントを編集者に追加しておく
        SPREADSHEET_KEY = '1GstVwQAGFUo2p5GoXHpQXeufYmctD7zEfLKGw-oG8r4'
        wb = gc.open_by_key(SPREADSHEET_KEY)
        ws = wb.get_worksheet(8)  #「シート2」を取得 # sheet1だと一番左のシートのみ。2つ目以降はsheet2にしてもエラーになる
 
        # 「Google スプレッドシート」に出力
        ws.update_cell(1, 1, '企業名')   # A1
        ws.update_cell(1, 2, 'webサイトURL')     # B1
        ws.update_cell(1, 3, '電話番号')     # C1
        sleep(5)

        print('ok')


        # A2からE89まで順に出力
        for i in range(1,len(Name)):
            print(f'{i}番目スタート')
            ws.update_cell(i+1, 1, Name[i-1])   # A列（name）
            ws.update_cell(i+1, 2, Website[i-1])     # B列（websiteURL）
            ws.update_cell(i+1, 3, Phone[i-1])     # C列（phone）

            # スプレッドシートの書き込みは1件/sなので一旦スリープさせる
            sleep(10)

        # 処理が終わったらwindowを閉じる
        self.driver.close()
        self.driver.quit()

if __name__ == '__main__':
    driver = Scrayping()
    driver.act('https://jan2021.tems-system.com/exhiSearch/INW/jp/ExhiList')