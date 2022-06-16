import time, requests, openpyxl, json, execjs, os, urllib3
import traceback
from hashlib import md5
import pandas as pd
from urllib.parse import quote
from selenium.webdriver import Chrome, ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

urllib3.disable_warnings()


class MapSearch:
    user_agents = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/102.0.0.0 Safari/537.36'

    def __init__(self):
        self.session = requests.session()
        try:
            self.company_lst = self.read_for_company()
        except:
            print('读取公司名称表失败')
            return

        self.filename = f'./公司名称表.xlsx'
        self.workbook = openpyxl.load_workbook(self.filename)
        self.worksheet = self.workbook.active

        with open('./decrypt.js', 'r', encoding='utf-8') as f:
            js = f.read()

        self.ctx = execjs.compile(js)

    def login(self):
        option = ChromeOptions()
        option.add_experimental_option("excludeSwitches", ["enable-automation"])
        option.add_experimental_option('useAutomationExtension', False)

        driver = Chrome(executable_path=r'./chromedriver.exe', options=option)
        driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
          "source": """
            Object.defineProperty(navigator, 'webdriver', {
              get: () => undefined
            })
          """
        })
        driver.execute_cdp_cmd("Network.enable", {})
        driver.execute_cdp_cmd("Network.setExtraHTTPHeaders", {"headers": {"User-Agent": self.user_agents}})
        driver.get('https://ditu.amap.com/')

        element = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "loginbox"))
        )
        element.click()
        time.sleep(2)
        driver.find_element_by_class_name('login-btn').click()

        WebDriverWait(driver, 1000).until(
            EC.presence_of_element_located((By.XPATH, '//div[@class="user-info"]/a[@class="user-name ellipsis"]'))
        )
        print(f'您已成功登录')
        cookie_lst = driver.get_cookies()
        driver.close()

        for cookie_dict in cookie_lst:
            self.session.cookies.set(cookie_dict['name'], cookie_dict['value'])

    def search_poi(self, city_code, search_key):
        # 08e7afc7361c9c2d8c6f25408d7e5b4b
        params = {
            'key': '117630351f2bad74f71d218bf699d39c', 'keywords': search_key,
            'types': '', 'city': city_code, 'children': '',
            'offset': 20, 'page': 1, 'extensions': 'all'
        }
        search_set = set([i for i in search_key])
        retry_count = 0
        while retry_count < 3:
            try:
                search_r = requests.get('https://restapi.amap.com/v3/place/text', params=params, verify=False)
                if search_r.json()['status'] == '1':
                    data = search_r.json()['pois']
                    for i in data:
                        if i['adcode'][:4] != str(city_code)[:4]:
                            continue

                        if i['name'] == search_key:
                            return i, f"关键词为: {search_key}. 已找到: {i['name']}. 精准匹配到. 已收录"
                        if search_key in i['name'] or i['name'] in search_key:
                            return i, f"关键词为: {search_key}. 已找到: {i['name']}. 模糊匹配到. 已收录"
                        r_key_set = set([i for i in i['name']])
                        match_score = len(search_set & r_key_set)/len(search_set)
                        if match_score >= 0.95:
                            return i, f"关键词为: {search_key}. 已找到: {i['name']}. 匹配度为: {round(match_score, 2)}. 已收录"

                    return f"关键词为: {search_key}. 未找到"
                else:
                    return f"关键词为: {search_key}. 未找到"
            except:
                retry_count += 1
                time.sleep(10)
                continue

        else:
            return f"关键词为: {search_key}. 未找到"

    def add_favorite(self, city_code, search_result):
        params, msg = search_result
        url = 'https://ditu.amap.com/service/fav/addFav?'
        headers = {
            'accept': '*/*', 'origin': 'https://ditu.amap.com', 'accept-encoding': 'gzip, deflate, br',
            'accept-language': 'zh-CN,zh;q=0.9', 'amapuuid': '1e6ae450-95a7-4f92-84fe-a18fb9a184d8',
            'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
            'referer': f'https://ditu.amap.com/place/{params["id"]}',
            'x-csrf-token': self.session.cookies.get('x-csrf-token'), 'User-Agent': self.user_agents,
            'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="102", "Google Chrome";v="102"',
            'sec-ch-ua-mobile': '?0', 'sec-ch-ua-platform': '"Windows"', 'sec-fetch-dest': 'empty',
            'sec-fetch-mode': 'cors', 'sec-fetch-site': 'same-origin', 'x-requested-with': 'XMLHttpRequest'
        }
        latitude, longitude = params['location'].split(',')
        result = self.ctx.call('get_data', longitude, latitude, 20)

        item_id = md5(f"{result['x']}+{result['y']}+{params['name']}".encode()).hexdigest()

        data = {
            'data[0][id]': item_id,
            'data[0][data][item_id]': item_id,
            'data[0][data][custom_address]': params['address'],
            'data[0][data][poiid]': params['id'],
            'data[0][data][custom_name]': params['name'],
            'data[0][data][type]': 0,
            'data[0][data][address]': params['address'],
            'data[0][data][phone_numbers]': '',
            'data[0][data][comment]': '',
            'data[0][data][name]': params['name'],
            'data[0][data][point_x]': result['x'],
            'data[0][data][point_y]': result['y'],
            'data[0][data][top_time]': '',
            'data[0][data][city_code]': city_code,
            'data[0][data][custom_phone_numbers]':'',
            'data[0][data][city_name]': params['cityname'],
            'data[0][data][tag]': '',
            'data[0][type]': 101,
            'data[0][ver]': 'YqqIEwAAAAABAAAB'
        }

        data = '&'.join([f'{quote(k)}={quote(str(v))}' for k, v in data.items()])

        retry_count = 0
        while retry_count < 3:
            try:
                r = self.session.post(url, data=data, headers=headers, verify=False)
                if r.json()['status'] == 1:
                    return msg
                else:
                    return False
            except:
                retry_count += 1
                time.sleep(10)
                continue

        else:
            return False

    def read_for_company(self):
        df = pd.read_excel('./公司名称表.xlsx', engine='openpyxl')
        company_lst = df['公司名称'].values.tolist()
        return company_lst

    def verify_city_code(self, city_name):
        with open('./city_code.json', 'r', encoding='utf-8') as f:
            city_code_dict = json.load(f)

        if city_name in city_code_dict:
            return city_code_dict[city_name]

    def start(self):
        search_key = input('请输入您要查询的城市: ').strip()

        city_code = self.verify_city_code(search_key)
        if not city_code:
            print('您输入的城市不存在')
            return self.start()

        self.login()

        for idx, company in enumerate(self.company_lst):
            print(f'正在搜索{company}')
            search_result = self.search_poi(city_code, company)
            if isinstance(search_result, tuple):
                add_stats = self.add_favorite(city_code, search_result)
                if add_stats:
                    print(add_stats)
                    self.worksheet.cell(row=idx+2, column=2).value = '存在, 已收录'
                    time.sleep(5)
                else:
                    self.worksheet.cell(row=idx + 2, column=2).value = '不存在'
            else:
                print(search_result)
                self.worksheet.cell(row=idx+2, column=2).value = '不存在'

            self.workbook.save(self.filename)

        self.workbook.close()
        self.session.close()
        print('所有数据已全部查询完毕')


if __name__ == '__main__':
    try:
        MapSearch().start()
    except:
        print(traceback.format_exc())

    os.system('pause')
