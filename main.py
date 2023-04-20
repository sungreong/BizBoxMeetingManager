import sys
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QCalendarWidget,
    QLabel,
    QComboBox,
    QLineEdit,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QHBoxLayout,
    QVBoxLayout,
    QMessageBox,
)
from PyQt5.QtGui import QFont
from PyQt5.QtCore import QDate, Qt
from src.register_calender_cls import BizBoxCommon
import configparser
import os 
def read_config():
    config = configparser.ConfigParser(interpolation=None)
    script_path = os.path.dirname(os.path.abspath(__file__))
    config.read(os.path.join(script_path,"./info.ini"))
    my_config_parser_dict = {s: dict(config.items(s)) for s in config.sections()}
    return my_config_parser_dict

class Schedule(QWidget):
    def __init__(self):
        super().__init__()
        self.config = read_config()
        
        # 3. 캘린더를 선택하면 날짜가 보이는 텍스트 박스가 필요
        self.date_text = QLineEdit(self)
        self.date_text.setGeometry(300, 100, 100, 30)

        # 4. 시작 시간과 종료 시간을 적을 수 있는 텍스트 박스가 필요
        self.start_time_text = QLineEdit(self)
        self.start_time_text.setGeometry(300, 140, 100, 30)

        self.end_time_text = QLineEdit(self)
        self.end_time_text.setGeometry(300, 180, 100, 30)

        # QHBoxLayout로 QLineEdit과 QLabel을 가로로 배치
        date_hbox = QHBoxLayout()
        date_label = QLabel("날짜:")
        date_hbox.addWidget(date_label)
        date_hbox.addWidget(self.date_text)

        # QHBoxLayout로 QLineEdit과 QLabel을 가로로 배치
        content_hbox = QHBoxLayout()
        content_label = QLabel("제목:")
        content_hbox.addWidget(content_label)
        self.content_text = QLineEdit(self)
        self.content_text.setGeometry(300, 100, 100, 30)
        content_hbox.addWidget(self.content_text)
        
        attendees_hbox = QHBoxLayout()
        attendees_label = QLabel("참석자:")
        attendees_hbox.addWidget(attendees_label)
        self.attendees_text = QLineEdit(self)
        self.attendees_text.setGeometry(300, 100, 100, 30)
        attendees_hbox.addWidget(self.attendees_text)
        

        start_time_hbox = QHBoxLayout()
        start_time_label = QLabel("시작시간:")
        start_time_hbox.addWidget(start_time_label)
        start_time_hbox.addWidget(self.start_time_text)

        end_time_hbox = QHBoxLayout()
        end_time_label = QLabel("종료시간:")
        end_time_hbox.addWidget(end_time_label)
        end_time_hbox.addWidget(self.end_time_text)

        # 1. 캘린더가 들어가 있어야 함.
        self.calendar = QCalendarWidget(self)
        self.calendar.setGridVisible(True)
        self.calendar.setGeometry(20, 20, 250, 200)
        self.calendar.clicked[QDate].connect(self.showDate)

        # 2. 버튼이 2개 있어야 함 (조회와 등록)
        self.view_btn = QPushButton("조회", self)
        self.view_btn.setGeometry(300, 20, 100, 30)
        self.view_btn.clicked.connect(self.view_schedule)

        self.reg_btn = QPushButton("등록", self)
        self.reg_btn.setGeometry(300, 60, 100, 30)
        self.reg_btn.clicked.connect(self.reg_schedule)

        # 5. combo box로 대,중,소 회의실이 있어야 해
        self.room_combo = QComboBox(self)
        self.room_combo.addItems(["2F) 대회의실", "2F) 소회의실", "3F) 대회의실", "3F) 소회의실", "3F) 중회의실","Zoom1","Zoom2"])
        self.room_combo.setGeometry(300, 220, 100, 30)

        self.filter_combo = QComboBox(self)
        self.filter_combo.addItems(["전체", "2F) 대회의실", "2F) 소회의실", "3F) 대회의실", "3F) 소회의실", "3F) 중회의실","Zoom1 : agilesoda@agilesoda.ai","Zoom2 : agilesoda2@agilesoda.ai"])
        self.filter_combo.setGeometry(20, 460, 100, 30)
        self.filter_combo.activated[str].connect(self.filter_schedule)

        # QHBoxLayout로 QPushButton을 가로로 배치
        btn_hbox = QHBoxLayout()
        btn_hbox.addWidget(self.filter_combo)
        btn_hbox.addWidget(self.view_btn)
        btn_hbox.addWidget(self.reg_btn)

        # QVBoxLayout로 QHBoxLayout을 세로로 배치
        vbox = QVBoxLayout()
        vbox.addLayout(date_hbox)
        vbox.addLayout(start_time_hbox)
        vbox.addLayout(end_time_hbox)
        vbox.addLayout(content_hbox)
        vbox.addLayout(attendees_hbox)
        vbox.addWidget(self.room_combo)
        vbox.addLayout(btn_hbox)

        # 6. 회의실을 등록하면, 테이블에 추가 되는 기능
        self.table = QTableWidget(self)
        self.table.setGeometry(20, 240, 380, 200)
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["날짜", "시작시간", "종료시간", "회의실", "내용"])

        font = QFont()
        font.setBold(True)
        self.table.horizontalHeader().setFont(font)
        self.table.horizontalHeader().setStyleSheet("QHeaderView::section {background-color: #f7d358}")

        # 전체 레이아웃을 QVBoxLayout으로 설정
        layout = QVBoxLayout()
        layout.addWidget(self.calendar)
        layout.addLayout(vbox)
        layout.addWidget(self.table)

        self.setLayout(layout)
        self.bizbox = BizBoxCommon(user=self.config["USER"]["id"],pw=self.config["USER"]["pw"])
        self.open_bizbox()

    def open_bizbox(self):
        self.bizbox.connect_site()
        self.bizbox.login()

    def filter_schedule(self, text):
        if text == "전체":
            for row in range(self.table.rowCount()):
                self.table.setRowHidden(row, False)
        else:
            for row in range(self.table.rowCount()):
                if self.table.item(row, 3).text() == text:
                    self.table.setRowHidden(row, False)
                else:
                    self.table.setRowHidden(row, True)

    def showDate(self, date):
        # 3. 캘린더를 선택하면 날짜가 보이는 텍스트 박스가 필요
        self.date_text.setText(date.toString(Qt.ISODate))

    def view_schedule(self):
        # 조회 버튼 클릭 시 실행되는 함수
        # 스케줄을 조회하는 기능을 추가해주세요.
        self.table.setRowCount(0)
        self.table.clearContents()
        
        date = self.date_text.text()
        try :
            df_room = self.bizbox.get_meeting_room_info(date)
        except Exception as e:
            QMessageBox.warning(self, "오류", str(e))
        else : 
            for idx, row in df_room.iterrows():
                row_count = self.table.rowCount()
                self.table.insertRow(row_count)
                self.table.setItem(row_count, 0, QTableWidgetItem(date))
                self.table.setItem(row_count, 1, QTableWidgetItem(row["시작시간"]))
                self.table.setItem(row_count, 2, QTableWidgetItem(row["종료시간"]))
                self.table.setItem(row_count, 3, QTableWidgetItem(row["장소"]))
                self.table.setItem(row_count, 4, QTableWidgetItem(row["회의명"]))
            else :
                QMessageBox.information(self, "알림", f"날짜 : {date} 회의 스케줄이 조회되었습니다.")
    def reg_schedule(self):
        # 등록 버튼 클릭 시 실행되는 함수
        # 스케줄을 등록하는 기능을 추가해주세요.
        date = self.date_text.text()
        start_time = self.start_time_text.text()
        end_time = self.end_time_text.text()
        room = self.room_combo.currentText()
        content = self.content_text.text()
        attendee_list = self.attendees_text.text().split(",")
        try :
            self.bizbox.add_schedule(title=content,date=date, start_time=start_time, end_time=end_time, meeting_room=room,attendee_list=attendee_list,)
        except Exception as e :
            QMessageBox.warning(self, "오류", str(e))
        else :
            # 등록된 스케줄을 테이블에 추가합니다.
            row_count = self.table.rowCount()
            self.table.insertRow(row_count)
            self.table.setItem(row_count, 0, QTableWidgetItem(date))
            self.table.setItem(row_count, 1, QTableWidgetItem(start_time))
            self.table.setItem(row_count, 2, QTableWidgetItem(end_time))
            self.table.setItem(row_count, 3, QTableWidgetItem(room))
            self.table.setItem(row_count, 4, QTableWidgetItem(content))
            QMessageBox.information(self, "알림", f"회의 스케줄이 등록되었습니다.\n날짜 : {date}\n미팅룸 : {room}\n시간 : {start_time}~{end_time} ")
    def resizeEvent(self, event):
        # 화면 크기가 조정될 때 테이블의 크기를 조정합니다.
        self.table.setGeometry(20, 300, self.width() - 40, self.height() - 350)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    schedule = Schedule()
    schedule.setGeometry(100, 100, 600, 700)
    schedule.setWindowTitle("BizBox 회의 스케줄 관리 프로그램")
    schedule.show()
    sys.exit(app.exec_())
