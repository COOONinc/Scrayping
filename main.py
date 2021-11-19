import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials 
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
        tableElem =self.driver.find_element_by_xpath('//*[@id="main"]/div[2]/table/tbody/tr[3]/td/table/tbody')
        trs = tableElem.find_elements(By.TAG_NAME, "tr")

        # 企業情報格納リスト
        Name = []
        Website = []
        Email = []
        Phone = []
        Address = []

        for num in range(2,len(trs)+1):
            self.driver.get(url)

            # 各企業のURL遷移
            self.driver.find_element_by_xpath(f'//*[@id="main"]/div[2]/table/tbody/tr[3]/td/table/tbody/tr[{num}]/td[1]/a').click()

            # 別タブ遷移型なのでwindowを変更
            handle_array = self.driver.window_handles # 遷移したタブを配列に格納

            self.driver.switch_to.window(handle_array[num-1]) #末尾に追加されていくので-1の値を取得
            sleep(5)

            page_source = self.driver.page_source
            soup = BeautifulSoup(page_source, 'html.parser')

            name = soup.find('h1')
            website = soup.select_one('div#exhibitor_details_website a')
            email = soup.select_one('div#exhibitor_details_email a')
            phone = soup.select_one('div#exhibitor_details_phone a')
            address = soup.select_one('div#exhibitor_details_address p')

            # 値が入っていない企業には空文字をセットする
            if(name is not None):
                Name.append(name.text)
            else:
                Name.append('')
            if(website is not None):
                Website.append(website.text)
            else:
                Website.append('')
            if(email is not None):
                Email.append(email.text)
            else:
                Email.append('')
            if(phone is not None):
                Phone.append(phone.text)
            else:
                Phone.append('')
            if(address is not None):
                Address.append(address.text)
            else:
                Address.append('')
 
            print(f'現在:{num}社目 / 残り{89-num}社')

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
        ws = wb.sheet1  # 一番左の「シート1」を取得
 
        # 「Google スプレッドシート」に出力
        ws.update_cell(1, 1, '企業名')   # A1
        ws.update_cell(1, 2, 'webサイトURL')     # B1
        ws.update_cell(1, 3, 'E-mail')     # C1
        ws.update_cell(1, 4, '電話番号')  # D1
        ws.update_cell(1, 5, '住所')  # E1
        sleep(5)

        print('ok')


        # A2からE89まで順に出力
        for i in range(1,len(Name)):
            print(f'{i}番目スタート')
            ws.update_cell(i+1, 1, Name[i-1])   # A列（name）
            ws.update_cell(i+1, 2, Website[i-1])     # B列（websiteURL）
            ws.update_cell(i+1, 3, Email[i-1])     # C列（email）
            ws.update_cell(i+1, 4, Phone[i-1])  # D列（phone）
            ws.update_cell(i+1, 5, Address[i-1])  # E列（address）

            # スプレッドシートの書き込みは1件/sなので一旦スリープさせる
            sleep(10)

        # 処理が終わったらwindowを閉じる
        self.driver.close()
        self.driver.quit()

if __name__ == '__main__':
    driver = Scrayping()
    driver.act('https://wsew2021-osaka.tems-system.com/ExhiSearch/WSEW/jp/ExhiList?_ga=2.109033221.623676628.1636432137-547389901.1636432137')