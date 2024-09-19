# -*- coding: utf-8 -*-
import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QWidget, QSlider, QLabel, QVBoxLayout, QHBoxLayout, QMessageBox, QPushButton, QSpinBox, QSizePolicy, QDesktopWidget, QFileDialog, QDialog, QLineEdit
)
from PyQt5.QtCore import QTimer, QTime, Qt, QDate, QUrl
from PyQt5.QtGui import QIcon, QFont
from PyQt5.QtMultimedia import QMediaPlayer, QMediaContent
from setting_class import ClassSettings, NoticeEditor  # NoticeSettings renamed to NoticeEditor
import json

class App(QWidget):
    def __init__(self):
        super().__init__()
        self.settings = {}
        self.load_settings()
        self.schedule = None 
        self.notice = None
        self.sound_played = False  # 플래그 추가

        # 소리 효과 초기화
        self.media_player = QMediaPlayer()
        sound_file = self.settings['sound_file_path']
        if os.path.exists(sound_file):
            self.media_player.setMedia(QMediaContent(QUrl.fromLocalFile(sound_file)))
            self.media_player.setVolume(50)  # 볼륨 설정 (0 ~ 100)
        else:
            print(f"소리 파일이 존재하지 않습니다: {sound_file}")

        self.current_date = QDate.currentDate()
        current_weekday_number = self.current_date.dayOfWeek()
        weekday_names = ["월요일", "화요일", "수요일", "목요일", "금요일", "토요일", "일요일"]
        self.current_weekday_name = weekday_names[current_weekday_number-1]
        
        # ClassSettings 객체 초기화 및 데이터 로드
        self.class_settings = ClassSettings(parent_setting=self.settings, load_only=True, current_weekday_name=self.current_weekday_name)  # UI를 표시하지 않도록 인자 추가
        self.period_times, self.weekly_timetable = self.class_settings.load_class_schedule()

        self.period_times = self.class_settings.period_times
        self.weekly_timetable = self.class_settings.weekly_timetable
        self.notices = self.class_settings.notices
        self.teacher_message = self.class_settings.teacher_message
        
        self.initUI()

    def closeEvent(self, event):
        reply = QMessageBox.question(self, '확인',
            "프로그램을 종료하시겠습니까?", QMessageBox.Yes |
            QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            self.closeAllWindows()
            event.accept()
        else:
            event.ignore()

    def closeAllWindows(self):
        # ClassSettings 창 닫기
        if self.class_settings:
            self.class_settings.close()
            # ClassSettings에서 관리하는 창들도 닫기
            if hasattr(self.class_settings, 'timetable_editor') and self.class_settings.timetable_editor:
                self.class_settings.timetable_editor.close()
            if hasattr(self.class_settings, 'notice_editor') and self.class_settings.notice_editor:
                self.class_settings.notice_editor.close()
            if hasattr(self.class_settings, 'period_editor') and self.class_settings.period_editor:
                self.class_settings.period_editor.close()
        
        # 추가적인 창이 있다면 여기에 닫는 코드를 추가하세요

        # 메인 애플리케이션 종료
        QApplication.quit()

    def load_settings(self):
        if not os.path.exists('setting.json'):
            self.settings = {
                "excel_file_path": "",
                "sound_file_path": "",
                "alarm_time_before": 1.0  # 기본값을 1분으로 설정
            }
            self.save_settings()
        else:
            with open('setting.json', 'r', encoding='utf-8') as f:
                self.settings = json.load(f)

    def save_settings(self):
        with open('setting.json', 'w', encoding='utf-8') as f:
            json.dump(self.settings, f, ensure_ascii=False, indent=4)

    def initUI(self):
        self.setStyleSheet("background-color: #E6E6FA;")  # 연한 보라색 배경
        main_layout = QVBoxLayout()
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # Title
        title_text = f"{self.current_date.toString('yyyy.MM.dd')} {self.current_weekday_name}"
        title_label = QLabel(title_text)
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("color: #4B0082; font-size: 80px; font-weight: bold;")
        main_layout.addWidget(title_label)

        # Time Display
        self.time_label = QLabel(self)
        self.time_label.setAlignment(Qt.AlignCenter)
        self.time_label.setStyleSheet("color: #4B0082; font-size: 180px; font-weight: bold; margin-bottom: 20px;")
        main_layout.addWidget(self.time_label)

        # Break Time Layout
        break_layout = QHBoxLayout()
        break_label_desc = QLabel("남은 쉬는 시간:")
        break_label_desc.setStyleSheet("color: #4B0082; font-size: 80px; font-weight: bold;")
        break_layout.addWidget(break_label_desc)

        self.break_label = QLabel(self)
        self.break_label.setStyleSheet("color: #32CD32; font-size: 80px; font-weight: bold;")
        break_layout.addWidget(self.break_label)
        break_layout.addStretch(1)

        main_layout.addLayout(break_layout)

        # Next Class Layout
        next_class_layout = QHBoxLayout()
        next_class_label_desc = QLabel("다음 수업:")
        next_class_label_desc.setStyleSheet("color: #4B0082; font-size: 80px; font-weight: bold;")
        next_class_layout.addWidget(next_class_label_desc)

        self.next_class_label = QLabel(self)
        self.next_class_label.setStyleSheet("color: #FF6347; font-size: 80px; font-weight: bold;")
        next_class_layout.addWidget(self.next_class_label)
        next_class_layout.addStretch(1)

        main_layout.addLayout(next_class_layout)

        # Notice Layout
        notice_layout = QVBoxLayout()
        self.notice_label = QLabel(self)
        self.notice_label.setStyleSheet("font-size:80px; background-color: #FFF5EE; padding: 50px; border-radius: 5px;")
        self.notice_label.setWordWrap(True)
        self.notice_label.setAlignment(Qt.AlignCenter)
        self.notice_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        main_layout.addWidget(self.notice_label)

        main_layout.addLayout(notice_layout)

        # Buttons Layout
        buttons_layout = QHBoxLayout()

        # Class Settings Button
        class_settings_button = QPushButton('수업 설정 열기', self)
        class_settings_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                padding: 10px;
                border-radius: 20px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        class_settings_button.clicked.connect(self.openClassSettings)
        buttons_layout.addWidget(class_settings_button)

        # Settings Button
        settings_button = QPushButton('프로그램 설정', self)
        settings_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                padding: 10px;
                border-radius: 20px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        settings_button.clicked.connect(self.openSettings)
        buttons_layout.addWidget(settings_button)

        main_layout.addLayout(buttons_layout)

        self.setLayout(main_layout)
        self.setWindowTitle("수업 알리미")
        self.setMinimumWidth(1300)  # 최소 너비 설정
        self.adjustSize()  # 컨텐츠에 맞게 크기 조절
        self.center()
        self.show()

        # Timer to update info every second
        timer = QTimer(self)
        timer.timeout.connect(self.update_info)
        timer.start(1000)
    
    def center(self):
        # 창을 화면의 중앙에 위치시키는 메서드
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def update_info(self):
        if self.current_weekday_name in ["토요일", "일요일"]:
            self.break_label.setText("쉬는 날")
            self.next_class_label.setText("쉬는 날")
            self.notice_label.setText("공지사항 없음")
            self.sound_played = False  # 플래그 초기화
            return

        currentTime = QTime.currentTime()
        hour = currentTime.hour()
        if hour > 12:
            hour -= 12
        label_time = f"{hour:02d}:{currentTime.minute():02d}:{currentTime.second():02d}"
        self.time_label.setText(label_time)

        try:
            notices, message = self.class_settings.load_notice()
        except AttributeError:
            notices, message = None, None

        remaining_time, next_class = self.calculate_remaining_time()
        try:
            next_class_name = self.weekly_timetable[self.current_weekday_name][next_class]
            self.break_label.setText(f"{remaining_time} 남음")
            self.next_class_label.setText(f"{next_class_name}")
            
            # 공지사항 업데이트
            if notices and next_class in notices and notices[next_class].strip():
                self.notice_label.setText(notices[next_class])
            # else:
            #     self.notice_label.setText("공지사항 없음")

        except KeyError:
            self.break_label.setText("없음")
            self.next_class_label.setText("없음")
            self.notice_label.setText("공지사항 없음")
            self.sound_played = False  # 플래그 초기화
            return

        # 남은 시간을 초 단위로 계산
        if remaining_time != "오늘 수업 없음" and remaining_time != "없음":
            mins, secs = map(int, remaining_time.replace("분 ", ":").replace("초", "").split(":"))
            total_seconds = mins * 60 + secs

            # 60초 이하일 때 레이블 색상을 빨간색으로 변경
            if total_seconds <= int(self.settings['alarm_time_before'] * 60):
                self.break_label.setStyleSheet("color: red; font-size: 80px; font-weight: bold;")
            else:
                self.break_label.setStyleSheet("color: #4B0082; font-size: 80px; font-weight: bold;")  # 기본 색상으로 변경

            # 알람 시간 설정 사용
            alarm_seconds = int(self.settings['alarm_time_before'] * 60)  # 분 단위를 초 단위로 변환
            if total_seconds <= alarm_seconds and not self.sound_played and self.media_player.media().isNull() == False:
                self.media_player.play()
                self.sound_played = True
            elif total_seconds > alarm_seconds:
                self.sound_played = False

    def openClassSettings(self):
        self.class_settings.initUI()
        self.class_settings.show()
        # 설정 창이 닫힐 때 데이터를 업데이트합니다.
        self.class_settings.closeEvent = self.updateSettingsData

    def updateSettingsData(self, event):
        self.period_times = self.class_settings.period_times
        self.weekly_timetable = self.class_settings.weekly_timetable
        self.notices = self.class_settings.notices
        self.teacher_message = self.class_settings.teacher_message
        event.accept()

    def load_class_times(self):
        # 이미 초기화된 class_settings를 사용
        self.period_times, self.weekly_timetable = self.class_settings.load_class_schedule()

    def calculate_remaining_time(self):
        currentTime = QTime.currentTime()
        smallest_diff = None
        next_class = None
        next_class_start_time = None

        for period, times in self.period_times.items():
            start_time = QTime.fromString(times[0], 'hh:mm')
            if currentTime < start_time:
                if smallest_diff is None or currentTime.secsTo(start_time) < smallest_diff:
                    smallest_diff = currentTime.secsTo(start_time)
                    next_class = period

        if smallest_diff is not None:
            remaining_time = f"{smallest_diff // 60}분 {smallest_diff % 60}초"
            return remaining_time, next_class
        return "오늘 수업 없음", next_class

    def openSettings(self):
        dialog = SettingsDialog(self.settings, self)
        if dialog.exec_():
            self.settings = dialog.get_settings()
            self.save_settings()
            self.load_class_settings()  # 클래스 설정을 다시 로드
            self.load_sound_effect()    # 소리 효과를 다시 로드

    def load_class_settings(self):
        # 엑셀 파일에서 수업 일정을 로드하는 코드
        excel_file = self.settings['excel_file_path']
        if os.path.exists(excel_file):
            # 여기에 엑셀 파일을 읽어 수업 일정을 설정하는 코드를 작성
            pass
        else:
            print(f"엑셀 파일을 찾을 수 없습니다: {excel_file}")

    def load_sound_effect(self):
        sound_file = self.settings['sound_file_path']
        if os.path.exists(sound_file):
            self.media_player = QMediaPlayer()
            self.media_player.setMedia(QMediaContent(QUrl.fromLocalFile(os.path.abspath(sound_file))))
            self.media_player.setVolume(self.settings.get('sound_volume', 50))  # 볼륨 설정 (0 ~ 100)
        else:
            self.media_player = None
            print(f"소리 파일을 찾을 수 없습니다: {sound_file}")

class SettingsDialog(QDialog):
    def __init__(self, settings, parent=None):
        super().__init__(parent)
        self.settings = settings.copy()
        self.media_player = QMediaPlayer()
        self.initUI()

    def initUI(self):
        self.setStyleSheet("""
            QDialog {
                background-color: #E6E6FA;
            }
            QLabel {
                color: #4B0082;
                font-size: 16px;
                font-weight: bold;
            }
            QLineEdit, QSpinBox {
                background-color: white;
                border: 1px solid #4B0082;
                border-radius: 5px;
                padding: 5px;
                font-size: 14px;
            }
            QPushButton {
                background-color: #4CAF50;
                color: white;
                padding: 8px;
                border-radius: 5px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)

        layout = QVBoxLayout()
        layout.setSpacing(20)
        layout.setContentsMargins(20, 20, 20, 20)

        # Title
        title_label = QLabel("프로그램 설정")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("font-size: 24px; margin-bottom: 10px;")
        layout.addWidget(title_label)

        # Excel file path
        excel_layout = QHBoxLayout()
        excel_label = QLabel("수업 시간표 파일 경로:")
        excel_layout.addWidget(excel_label)
        self.excel_path = QLineEdit(self.settings['excel_file_path'])
        excel_layout.addWidget(self.excel_path)
        excel_button = QPushButton("찾기")
        excel_button.clicked.connect(self.browse_excel)
        excel_layout.addWidget(excel_button)
        layout.addLayout(excel_layout)

        # Sound file path
        sound_layout = QHBoxLayout()
        sound_label = QLabel("알람 소리 파일 경로:")
        sound_layout.addWidget(sound_label)
        self.sound_path = QLineEdit(self.settings['sound_file_path'])
        sound_layout.addWidget(self.sound_path)
        sound_button = QPushButton("찾기")
        sound_button.clicked.connect(self.browse_sound)
        sound_layout.addWidget(sound_button)
        layout.addLayout(sound_layout)

        # Alarm time before
        alarm_layout = QHBoxLayout()
        alarm_label = QLabel("알람 시간:")
        alarm_layout.addWidget(alarm_label)
        
        self.alarm_minutes = QSpinBox()
        self.alarm_minutes.setRange(0, 10)  # 0분부터 10분까지
        self.alarm_minutes.setSuffix(" 분")
        alarm_layout.addWidget(self.alarm_minutes)
        
        self.alarm_seconds = QSpinBox()
        self.alarm_seconds.setRange(0, 50)  # 0초부터 50초까지
        self.alarm_seconds.setSingleStep(10)  # 10초 단위로 증가
        self.alarm_seconds.setSuffix(" 초")
        alarm_layout.addWidget(self.alarm_seconds)
        
        # 기존 설정값 적용
        total_seconds = int(self.settings['alarm_time_before'] * 60)
        self.alarm_minutes.setValue(total_seconds // 60)
        self.alarm_seconds.setValue((total_seconds % 60) // 10 * 10)  # 10초 단위로 반올림
        
        layout.addLayout(alarm_layout)

        volume_layout = QHBoxLayout()
        volume_label = QLabel("알람 소리 크기:")
        volume_layout.addWidget(volume_label)
        
        self.volume_slider = QSlider(Qt.Horizontal)
        self.volume_slider.setRange(0, 100)
        self.volume_slider.setValue(self.settings.get('sound_volume', 50))
        self.volume_slider.valueChanged.connect(self.update_volume)
        volume_layout.addWidget(self.volume_slider)
        
        self.volume_value_label = QLabel(f"{self.volume_slider.value()}%")
        volume_layout.addWidget(self.volume_value_label)
        
        self.preview_button = QPushButton("미리 듣기")
        self.preview_button.clicked.connect(self.preview_sound)
        volume_layout.addWidget(self.preview_button)
        
        layout.addLayout(volume_layout)

        # Buttons
        button_layout = QHBoxLayout()
        ok_button = QPushButton("확인")
        ok_button.clicked.connect(self.accept)
        cancel_button = QPushButton("취소")
        cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)

        self.setLayout(layout)
        self.setWindowTitle('설정')
        self.setMinimumWidth(450)
        self.setMinimumHeight(300)
        self.adjustSize()

    def browse_excel(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "수업 시간표 파일 선택", "", "Excel Files (*.xlsx *.xls)")
        if file_name:
            self.excel_path.setText(file_name)

    def browse_sound(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "소리 파일 선택", "", "Sound Files (*.mp3 *.wav)")
        if file_name:
            self.sound_path.setText(file_name)

    def update_volume(self, value):
        self.volume_value_label.setText(f"{value}%")
        self.media_player.setVolume(value)

    def preview_sound(self):
        sound_file = self.sound_path.text()
        if os.path.exists(sound_file):
            self.media_player.setMedia(QMediaContent(QUrl.fromLocalFile(sound_file)))
            self.media_player.play()
        else:
            print(f"소리 파일을 찾을 수 없습니다: {sound_file}")

    def get_settings(self):
        self.settings['excel_file_path'] = self.excel_path.text()
        self.settings['sound_file_path'] = self.sound_path.text()
        total_seconds = self.alarm_minutes.value() * 60 + self.alarm_seconds.value()
        self.settings['alarm_time_before'] = total_seconds / 60  # 분 단위로 저장
        self.settings['sound_volume'] = self.volume_slider.value()
        return self.settings

# Boilerplate code to run the application
if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon('./source/alarmi_icon.ico'))
    font = QFont("나눔고딕", 10)  # 폰트 이름과 크기
    app.setFont(font)
    ex = App()
    sys.exit(app.exec_())