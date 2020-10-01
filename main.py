from time import sleep
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import random
import json
import requests
import os

min_temperature = 36.0
max_temperature = 36.8

# ドライバーの設定
# driver_path = 'driver\\chromedriver.exe'
driver_path = '/app/.chromedriver/bin/chromedriver'

options = Options()
options.add_argument('--disable-gpu')
options.add_argument('--disable-extensions')
options.add_argument('--proxy-server="direct://"')
options.add_argument('--proxy-bypass-list=*')
options.add_argument('--start-maximized')
options.add_argument('--headless')

driver = webdriver.Chrome(executable_path=driver_path, chrome_options=options)
driver.implicitly_wait(3)
url = 'https://login.microsoftonline.com/common/oauth2/authorize?response_type=id_token&client_id=5e3ce6c0-2b1f-4285-8d4b-75ee78787346&redirect_uri=https%3A%2F%2Fteams.microsoft.com%2Fgo&state=a38d4d10-b098-4a8e-ab23-d4149b93d27d&client-request-id=e424444d-a4e6-4cfe-8d03-26e9f05699a3&x-client-SKU=Js&x-client-Ver=1.0.9&nonce=73c6e26e-479b-4760-b7d9-de262416763d&domain_hint=&sso_reload=true'
driver.get(url)
driver.maximize_window()


def auto_point_call(address, password, name, i):
    # メールアドレス
    elem = driver.find_element_by_id('i0116')
    elem.clear()
    elem.send_keys(address)
    sleep(2)
    elem = driver.find_element_by_id('idSIButton9')
    elem.click()

    # パスワード
    sleep(3)
    elem = driver.find_element_by_id('i0118')
    elem.send_keys(password)
    # ロードされるのを待つ
    sleep(4)
    elem = driver.find_element_by_id('idSIButton9')
    elem.click()
    if i == 0:
        # ロードされるのを待つ
        sleep(4)
        elem = driver.find_element_by_id('idBtn_Back')
        elem.click()

    # 代わりにwebアプリを使用する
    sleep(3)
    elem = driver.find_element_by_xpath(
        '//*[@id="download-desktop-page"]/div/a')
    elem.click()
    # チームタイル
    sleep(2)
    elem = driver.find_element_by_xpath(
        '//*[@id="favorite-teams-panel"]/div/div[1]/div[2]/div[2]/div/ng-include/div/div')
    elem.click()
    # タブの選択
    sleep(2)
    elem = driver.find_element_by_xpath(
        '//*[@id="tab::1066f8af-f34a-4ad5-bd26-4b4b5facbce2"]/div/a')
    elem.click()

    # ブラウザマーク
    elem = driver.find_element_by_css_selector(
        '.app-svg.icons-website'
    )
    elem.click()

    # 入力欄をクリック
    sleep(2)
    driver.switch_to.window(driver.window_handles[-1])
    elem = driver.find_element_by_css_selector(
        '.office-form-question-textbox.office-form-textfield-input.form-control.border-no-radius')
    elem.click()
    # 名前入力
    elem.send_keys(name)
    # 学年の選択
    elem = driver.find_element_by_xpath("//input[@value='3年']")
    elem.click()
    sleep(1)
    # 学科の選択
    sleep(1)
    course_of_study_dict = {
        'm': "//input[@value='M科']",
        'e': "//input[@value='E科']",
        'j': "//input[@value='J科']",
        'k': "//input[@value='K科']",
        'c': "//input[@value='C科']"

    }
    elem = driver.find_element_by_xpath(course_of_study_dict[address[0]])
    elem.click()
    sleep(1)
    # 体調の選択
    elem = driver.find_element_by_xpath("//input[@value='体調に異常はない']")
    elem.click()

    # 体温の選択
    body_temperature = round(random.uniform(
        min_temperature, max_temperature), 1)
    body_temperature_dict = {
        36.0: '7',
        36.1: '8',
        36.2: '9',
        36.3: '10',
        36.4: '11',
        36.5: '12',
        36.6: '13',
        36.7: '14',
        36.8: '15'
    }

    elem = driver.find_element_by_css_selector(
        '.ms-Icon.ms-Icon--ChevronDown.select-placeholder-arrow.forms-icon-size14x14')
    elem.click()
    xpath = '//*[@id="e9"]/ul/li[' + \
            body_temperature_dict[body_temperature] + ']/div/span'
    elem = driver.find_element_by_xpath(xpath)
    elem.click()
    sleep(1)
    # 送信
    elem = driver.find_element_by_css_selector(
        '.__submit-button__.office-form-bottom-button.office-form-theme-button.button-control.light-background-button')
    # elem.click()
    sleep(4)
    # サインアウト
    driver.switch_to.window(driver.window_handles[0])
    elem = driver.find_element_by_xpath(
        '//*[@id="personDropdown"]/div/div/div/profile-picture/img')
    elem.click()
    sleep(2)
    elem = driver.find_element_by_xpath('//*[@id="logout-button"]')
    elem.click()
    sleep(10)


def send_message():
    line_url = "https://notify-api.line.me/api/notify"
    headers = {'Authorization': 'Bearer ' +
                                'R5zwfjLSPLol0ldA2bY0vhWQgohr1I5NugA2LiQaTG7'}
    payload = {'message': '全員の送信が完了.安心してね！'}
    requests.post(line_url, headers=headers, params=payload, )


def main():
    file = open("userinfo.json",
                "r", encoding="utf-8")
    opened_json = json.load(file)
    dict_len = len(opened_json["data"])
    for i in range(dict_len):
        address = opened_json["data"][i]['address']
        name = opened_json["data"][i]['name']
        password = os.environ[str(i)]
        auto_point_call(address=address, name=name, password=password, i=i)
        driver.find_element_by_xpath('//*[@id="otherTileText"]').click()

    driver.quit()
    send_message()
