from selenium import webdriver
import chromedriver_autoinstaller
from time import sleep
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

from selenium.webdriver.common.keys import Keys
import pandas as pd
import re
from selenium.webdriver.support.ui import Select

class BizBoxCommon(object):
    def __init__(self, user="##", pw="##"):
        self.user = user
        self.pw = pw

    def open(self):
        options = webdriver.ChromeOptions()
        options.add_argument("headless")
        options.add_argument("--window-size=1920,1000")
        options.add_argument("disable-gpu")
        path = chromedriver_autoinstaller.install()
        self.driver = webdriver.Chrome(path, options=options)

    def close(self):
        self.driver.close()
    
    def connect_site(self):
        self.open()
        print("사이트로 접속합니다.")
        self.driver.get("http://gw.agilesoda.ai/gw/uat/uia/egovLoginUsr.do")

    def login(self):
        search_box_id = self.driver.find_element_by_xpath('//*[@id="userId"]')
        search_box_passwords = self.driver.find_element_by_xpath('//*[@id="userPw"]')
        login_btn = self.driver.find_element_by_xpath(
            '//*[@id="login_b1_type"]/div[2]/div[2]/form/fieldset/div[2]/div'
        )

        print("로그인을 실시합니다")
        search_box_id.send_keys(self.user)
        search_box_passwords.send_keys(self.pw)
        login_btn.click()
        sleep(0.5)
        self.access_main()
        self.driver.find_element_by_xpath("//div[@class='fm_div fm_sc']").click()
    
        button = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//div[@id="1dep"]/span'))
        )
        button.click()
        
    def access_main(self):
        self.driver.switch_to_window(self.driver.window_handles[0])  # main page

    def get_meeting_room_info(self, date="YYYY-MM-DD"):
        try :
            
            self.driver.switch_to.default_content()
            # 링크 엘리먼트 찾기
            link = self.driver.find_element_by_id("302020000_anchor")
            link.click()
            # iframe_wrap 엘리먼트 찾기
            iframe_element = self.driver.find_element_by_tag_name("iframe")

            self.driver.switch_to.frame(iframe_element)

            """
            # 회의실 예약 정보 확인하기
            """

            date_element = self.driver.find_element_by_id("from_date")
            self.driver.execute_script(f"arguments[0].value = '{date}';", date_element)

            date_element = self.driver.find_element_by_id("to_date")
            self.driver.execute_script(f"arguments[0].value = '{date}';", date_element)

            # 버튼 요소를 찾습니다.
            button = self.driver.find_element_by_id("searchButton")

            # 버튼을 클릭합니다.
            button.click()

            sleep(0.1)
            # 페이지에서 k-pager-numbers 클래스를 가진 엘리먼트 가져오기
            page_numbers = self.driver.find_element_by_class_name("k-pager-numbers")

            # 페이지 숫자가 있는 엘리먼트를 가져오기
            page_links = page_numbers.find_elements_by_tag_name("a")

            # 각 링크의 data-page 속성 값 가져오기
            max_page_number = int(max([page_link.get_attribute("data-page") for page_link in page_links]))
        except Exception as e : 
            raise Exception(e)
        # 데이터를 저장할 빈 리스트 생성
        data = []
        for page_number in range(1, max_page_number + 1):

            # 페이지 숫자를 클릭하고자 하는 요소를 찾습니다.
            page_number = self.driver.find_element_by_xpath(f"//a[@data-page='{page_number}']")

            # 요소를 클릭합니다.
            page_number.click()

            # 테이블의 각 행을 순회하며 데이터 추출
            for row in self.driver.find_elements_by_css_selector("table[role='grid'] tr"):
                # 각 셀의 값을 추출하여 리스트에 추가
                cols = row.find_elements_by_css_selector("td")
                cols = [col.text for col in cols]
                data.append(cols)
            else:
                sleep(0.001)
        else:
            # 추출된 데이터를 데이터프레임으로 변환
            df = pd.DataFrame(data, columns=["시간", "장소", "회의명", "상태", "예약자"])
            df = df[df.isnull().sum(axis=1) != 5]
            df = df[df["상태"] == "예약완료"]

        def get_start_end_datetime(date_str):
            def get_date_pattern(date_str):
                date_match = re.match(r"(\d{4}\.\d{2}\.\d{2})\((\w{1,3})\)\s+(\d{2}:\d{2})", date_str)
                date, day, time = date_match.group(1), date_match.group(2), date_match.group(3)
                return date, time

            start_time_str, end_time_str = date_str.split(" - ")
            start_date, start_time = get_date_pattern(start_time_str)
            end_date, end_time = get_date_pattern(end_time_str)
            return [start_date, end_date, start_time, end_time]

        df_room = df[~df["장소"].str.contains("Zoom")]
        df_zoom = df[df["장소"].str.contains("Zoom")]

        df_room[["시작날짜", "종료날짜", "시작시간", "종료시간"]] = df_room["시간"].apply(get_start_end_datetime).apply(pd.Series)
        df_room.drop(columns=["시간"], inplace=True)
        df_room = df_room.sort_values(by=["장소", "시작날짜", "시작시간"]).reset_index(drop=True)
        ROOM_LIST = ["2F) 대회의실", "2F) 소회의실", "3F) 대회의실", "3F) 소회의실", "3F) 중회의실"]
        print(df_room.query(f'장소 == "{ROOM_LIST[0]}"').to_markdown())
        return df_room

    def add_schedule(self, title="제목", date="2021-01-01", start_time="09:00", end_time="10:00", meeting_room="2F) 대회의실") :
                    
        # iFrame에서 기본 페이지로 돌아가기
        self.driver.switch_to.default_content()
        # 링크 엘리먼트 찾기
        link = self.driver.find_element_by_id("302020000_anchor")
        link.click()

        # iframe_wrap 엘리먼트 찾기
        iframe_element = self.driver.find_element_by_tag_name("iframe")

        self.driver.switch_to.frame(iframe_element)

        # 버튼 클릭
        button_element = self.driver.find_element_by_id("registRes")
        button_element.click()


        search_title_id = self.driver.find_element_by_xpath('//*[@id="res_reserve_name"]')
        search_title_id.send_keys(title)
        date_element = self.driver.find_element_by_id("from_date")
        date_element.send_keys(date)
        date_element.send_keys(Keys.ENTER)

        # 변경된 값을 확인
        changed_value = date_element.get_attribute("value")

        # 자바스크립트 코드 실행
        self.driver.execute_script(f'document.getElementById("from_date").value = "{date}";')
        self.driver.execute_script(f'document.getElementById("to_date").value = "{date}";')

        # Select 객체 생성
        select_box = Select(self.driver.find_element_by_id("sc_combox_2"))
        # 24:00 선택
        select_box.select_by_value(start_time)

        select_box = Select(self.driver.find_element_by_id("sc_combox_3"))
        # 24:00 선택
        select_box.select_by_value(end_time)
        sleep(1)
        # resource_search 접근
        button = self.driver.find_element_by_xpath("//button[@onclick='javascript:resourceSearch(1)']")  # 버튼을 찾습니다.
        button.click()

        sleep(1)
        element = self.driver.find_element_by_xpath(f'//span[contains(text(), "{meeting_room}")]')
        element.click()

        # 버튼 element 탐색
        button = self.driver.find_element_by_xpath('//div[@id="resource_search"]//input[@value="확인"]')

        # 버튼 클릭
        button.click()
        ##########################
        # 추가 인원 등록 
        # 사용자 추가
        # self.driver.switch_to_window(self.driver.window_handles[0])  # search page
        # iframe_element = self.driver.find_element_by_tag_name("iframe")

        # self.driver.switch_to.frame(iframe_element)
        # button = WebDriverWait(self.driver, 10).until(
        #     EC.presence_of_element_located((By.XPATH, "//button[contains(text(), '선택')]"))
        # )

        # 버튼을 클릭합니다.
        # button.click()
        # # 해당 input 태그를 찾습니다.
        # self.driver.switch_to_window(self.driver.window_handles[-1])  # search page

        # input_elem = self.driver.find_element_by_id("text_input")

        # # input 태그에 값을 입력합니다.
        # input_elem.send_keys("")
        # input_elem.send_keys("송유한")

        # # 버튼을 클릭합니다.

        # link_elem = self.driver.find_element_by_css_selector("a.btn_sear.btn_search")
        # link_elem.click()
        # # 체크 박스를 찾아 클릭합니다.
        # checkbox_elem = self.driver.find_element_by_css_selector("label[for='inp_chk_indexOf3']")
        # checkbox_elem.click()


        # 버튼 element 탐색
        # button = self.driver.find_element_by_xpath('//div[@id="btn_cen.pt12"]//input[@value="저장"]')
        # # 해당 버튼을 찾습니다.
        # button = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, "//input[@id='btn_save']")))
        # # 버튼 클릭
        # button.click()
        ##########################

        self.driver.switch_to_window(self.driver.window_handles[0])  # main page
        iframe_element = self.driver.find_element_by_tag_name("iframe")
        self.driver.switch_to.frame(iframe_element)
        self.driver.execute_script("window.scrollBy(0, 800);")

        # 해당 버튼을 찾습니다.
        button = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, "//input[@id='save_regist']")))
        button.click()
