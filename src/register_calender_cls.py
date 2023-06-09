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
from PyQt5.QtWidgets import QApplication, QProgressBar


class BizBoxCommon(object):
    def __init__(self, user="##", pw="##"):
        self.user = user
        self.pw = pw
        self.click_count = 0

    def add_progressbar(self, progress_bar: QProgressBar):
        self.progress_bar = progress_bar

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
        try:
            search_box_id.send_keys(self.user)
            search_box_passwords.send_keys(self.pw)
            login_btn.click()
            sleep(0.5)
            self.access_main()
            self.driver.find_element_by_xpath("//div[@class='fm_div fm_sc']").click()

            if self.click_count % 2 == 0:
                button = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//div[@id="1dep"]/span'))
                )
                button.click()
                self.click_count += 1
            else:
                pass
        except Exception as e:
            print("로그인 정보가 올바르지 않은 경우, info.ini 파일생성 후, id와 pw가 잘 설정되어있는 지 확인해주세요.")
            import os
            from pathlib import Path

            folder_path = str(Path(__file__).parent.parent)
            os.startfile(folder_path)
            raise Exception(e)

    def access_main(self):
        self.driver.switch_to_window(self.driver.window_handles[0])  # main page

    def access_sub(self):
        self.driver.get("http://gw.agilesoda.ai/gw/userMain.do")
        self.driver.find_element_by_xpath("//div[@class='fm_div fm_sc']").click()
        button = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//div[@id="1dep"]/span'))
        )
        button.click()
        self.click_count += 1

    def add_progressbar_event(self, value):
        self.progress_bar.setValue(value)
        QApplication.processEvents()  # GUI 업데이트

    def get_meeting_room_info(self, date="YYYY-MM-DD"):
        self.progress_bar.show()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(10)
        self.progress_bar.setValue(0)

        try:
            self.driver.switch_to.default_content()
            # 링크 엘리먼트 찾기
            link = self.driver.find_element_by_id("302020000_anchor")
            link.click()
            self.add_progressbar_event(1)

            # iframe_wrap 엘리먼트 찾기
            iframe_element = self.driver.find_element_by_tag_name("iframe")

            self.driver.switch_to.frame(iframe_element)
            self.add_progressbar_event(2)
            sleep(0.5)
            """
            # 회의실 예약 정보 확인하기
            """
            print(f"조회할 회의실 날짜 선택 {date}")
            date_element = self.driver.find_element_by_id("from_date")
            self.driver.execute_script(f"arguments[0].value = '{date}';", date_element)

            date_element = self.driver.find_element_by_id("to_date")
            self.driver.execute_script(f"arguments[0].value = '{date}';", date_element)
            sleep(0.5)
            # 버튼 요소를 찾습니다.
            print(f"조회할 회의실 날짜 클릭 {date}")
            button = self.driver.find_element_by_id("searchButton")

            # 버튼을 클릭합니다.
            button.click()
            self.add_progressbar_event(3)

            sleep(1)
            print(f"조회할 페이지 선택 ")
            # 페이지에서 k-pager-numbers 클래스를 가진 엘리먼트 가져오기
            page_numbers = self.driver.find_element_by_class_name("k-pager-numbers")

            # 페이지 숫자가 있는 엘리먼트를 가져오기
            page_links = page_numbers.find_elements_by_tag_name("a")
            # 각 링크의 data-page 속성 값 가져오기
            if len(page_links) == 0:
                max_page_number = 1
            else:
                max_page_number = int(max([page_link.get_attribute("data-page") for page_link in page_links]))
        except Exception as e:
            print(e)
            raise Exception(e)
        # 데이터를 저장할 빈 리스트 생성
        print(f"페이지별로 데이터 조회")
        data = []
        self.add_progressbar_event(4)
        for page_number in range(1, max_page_number + 1):
            print(f"페이지별로 데이터 조회... {page_number}/{max_page_number}")

            # 페이지 숫자를 클릭하고자 하는 요소를 찾습니다.
            page_number_button = self.driver.find_element_by_xpath(f"//a[@data-page='{page_number}']")

            # 요소를 클릭합니다.
            page_number_button.click()
            sleep(0.01)
            # 테이블의 각 행을 순회하며 데이터 추출
            for row in self.driver.find_elements_by_css_selector("table[role='grid'] tr"):
                # 각 셀의 값을 추출하여 리스트에 추가
                cols = row.find_elements_by_css_selector("td")
                cols = [col.text for col in cols]
                data.append(cols)
            else:
                self.add_progressbar_event(int(4 + page_number))
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

        df_room = df
        # df_room = df[~df["장소"].str.contains("Zoom")]
        # df_zoom = df[df["장소"].str.contains("Zoom")]

        df_room[["시작날짜", "종료날짜", "시작시간", "종료시간"]] = df_room["시간"].apply(get_start_end_datetime).apply(pd.Series)
        df_room.drop(columns=["시간"], inplace=True)
        df_room = df_room.sort_values(by=["장소", "시작날짜", "시작시간"]).reset_index(drop=True)
        # ROOM_LIST = ["2F) 대회의실", "2F) 소회의실", "3F) 대회의실", "3F) 소회의실", "3F) 중회의실"]
        # print(df_room.query(f'장소 == "{ROOM_LIST[0]}"').to_markdown())
        self.add_progressbar_event(10)
        return df_room

    def add_schedule(
        self,
        title="제목",
        date="2021-01-01",
        start_time="09:00",
        end_time="10:00",
        meeting_room="2F) 대회의실",
        attendee_list=[],
    ):
        # iFrame에서 기본 페이지로 돌아가기
        self.progress_bar.show()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(12)
        self.progress_bar.setValue(0)
        self.access_sub()
        sleep(1.0)
        self.add_progressbar_event(1)
        # 링크 엘리먼트 찾기

        link = self.driver.find_element_by_id("302020000_anchor")
        link.click()
        sleep(1.0)
        self.add_progressbar_event(2)

        # iframe_wrap 엘리먼트 찾기
        iframe_element = self.driver.find_element_by_tag_name("iframe")

        self.driver.switch_to.frame(iframe_element)
        sleep(0.5)
        print(f"등륵 버튼 클릭")
        self.add_progressbar_event(3)
        # 버튼 클릭
        button_element = self.driver.find_element_by_id("registRes")
        button_element.click()
        sleep(0.5)
        self.add_progressbar_event(4)
        print(f"회의실 멍 셜정")
        search_title_id = self.driver.find_element_by_xpath('//*[@id="res_reserve_name"]')
        search_title_id.send_keys(title)
        date_element = self.driver.find_element_by_id("from_date")
        date_element.send_keys(date)
        date_element.send_keys(Keys.ENTER)
        sleep(0.5)
        self.add_progressbar_event(5)
        # 변경된 값을 확인
        changed_value = date_element.get_attribute("value")

        # 자바스크립트 코드 실행
        print(f"회의실 기간 설정")
        self.driver.execute_script(f'document.getElementById("from_date").value = "{date}";')
        self.driver.execute_script(f'document.getElementById("to_date").value = "{date}";')

        print(f"회의실 시간 설정")
        select_box = Select(self.driver.find_element_by_id("sc_combox_2"))
        # 24:00 선택
        select_box.select_by_value(start_time)

        select_box = Select(self.driver.find_element_by_id("sc_combox_3"))
        # 24:00 선택
        select_box.select_by_value(end_time)
        sleep(1)
        self.add_progressbar_event(6)
        # resource_search 접근
        print(f"회의실 자원 설정")
        button = self.driver.find_element_by_xpath("//button[@onclick='javascript:resourceSearch(1)']")  # 버튼을 찾습니다.
        button.click()

        sleep(1)
        self.add_progressbar_event(7)
        element = self.driver.find_element_by_xpath(f'//span[contains(text(), "{meeting_room}")]')
        element.click()
        self.add_progressbar_event(8)
        # 버튼 element 탐색
        button = self.driver.find_element_by_xpath('//div[@id="resource_search"]//input[@value="확인"]')

        # 버튼 클릭
        button.click()
        self.add_progressbar_event(9)
        ##########################
        # 추가 인원 등록
        # 사용자 추가

        print(f"사용자 추가 설정")
        self.driver.switch_to_window(self.driver.window_handles[0])  # search page
        iframe_element = self.driver.find_element_by_tag_name("iframe")

        self.driver.switch_to.frame(iframe_element)
        button = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//button[contains(text(), '선택')]"))
        )
        # 버튼을 클릭합니다.
        button.click()
        sleep(1)
        self.add_progressbar_event(10)
        if attendee_list == []:
            pass
        else:
            # 해당 input 태그를 찾습니다.
            self.driver.switch_to_window(self.driver.window_handles[-1])  # search page

            input_elem = self.driver.find_element_by_id("text_input")
            # input 태그에 값을 입력합니다.
            for person in attendee_list:
                input_elem.clear()
                input_elem.send_keys(person)
                # 버튼을 클릭합니다.
                link_elem = self.driver.find_element_by_css_selector("a.btn_sear.btn_search")
                link_elem.click()
                # 체크 박스를 찾아 클릭합니다.
                # checkbox_elem = self.driver.find_element_by_css_selector("label[for='inp_chk_indexOf4']")
                sleep(0.5)
                element = self.driver.find_element_by_xpath(
                    f"//td[contains(text(), '{person}')]/..//input[@type='checkbox']"
                )
                self.driver.execute_script("arguments[0].scrollIntoView(true);", element)
                self.driver.execute_script("arguments[0].click();", element)
            print(f"사용자 확정")
            button = self.driver.find_element_by_id("btn_save")
            self.driver.execute_script("arguments[0].click();", button)
        self.add_progressbar_event(11)
        ##########################

        self.driver.switch_to_window(self.driver.window_handles[0])  # main page
        iframe_element = self.driver.find_element_by_tag_name("iframe")
        self.driver.switch_to.frame(iframe_element)
        self.driver.execute_script("window.scrollBy(0, 800);")
        self.add_progressbar_event(12)
        # 해당 버튼을 찾습니다.
        button = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//input[@id='save_regist']"))
        )
        button.click()
        self.progress_bar.hide()
