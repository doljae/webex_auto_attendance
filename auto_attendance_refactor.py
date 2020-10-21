#-*- coding:utf-8 -*-
import sys
from PyQt5.QtWidgets import *
import pandas as pd
import datetime
import glob

#UI파일 연결
#단, UI파일은 Python 코드 파일과 같은 디렉토리에 위치해야한다.
# form_class = uic.loadUiType("auto_attendance.ui")[0]

class WindowClass(QWidget):
    def __init__(self):
        super().__init__()
        self.setupUI()

    def setupUI(self):
        self.setGeometry(800, 200, 400, 300)
        self.setWindowTitle("Webex 자동 출결 기록 시스템 v2.1")

        self.label1 = QLabel()
        self.label3 = QLabel()
        self.label4 = QLabel()
        self.label2 = QLabel()

        self.label5=QLabel()

        self.pushButton1 = QPushButton("수강생명단.csv를 선택해주세요.")
        self.pushButton1.clicked.connect(self.pushButtonClicked1)

        self.pushButton4 = QPushButton("수업시작시간을 입력해주세요")
        self.pushButton4.clicked.connect(self.pushButtonClicked4)

        self.pushButton2 = QPushButton("유예 시간을 정해주세요")
        self.pushButton2.clicked.connect(self.pushButtonClicked2)

        self.pushButton3 = QPushButton("실행")
        self.pushButton3.clicked.connect(self.solution)

        layout = QVBoxLayout()
        layout.addWidget(self.pushButton1)
        layout.addWidget(self.label1)
        layout.addWidget(self.pushButton4)
        layout.addWidget(self.label4)
        layout.addWidget(self.pushButton2)
        layout.addWidget(self.label2)
        layout.addWidget(self.pushButton3)
        layout.addWidget(self.label3)
        self.label3.setText("Github repo: https://github.com/doljae/webex_auto_attendance")
        self.setLayout(layout)

    def solution(self):
        STUDENT_LIST_CSV = self.label1.text()
        LIMIT_ATTENDANCE_TIME=self.label4.text()
        AFFORDABLE_MINUTE = self.label2.text()

        path = "./text_list/*"
        file_list = glob.glob(path)
        file_list_txt = [file for file in file_list if file.endswith(".txt")]
        file_list_txt.sort()

        limit_datetime_object1 = datetime.datetime.strptime(LIMIT_ATTENDANCE_TIME, "%H:%M")
        limit_datetime_object = limit_datetime_object1 + datetime.timedelta(minutes=int(AFFORDABLE_MINUTE))
        df1 = pd.read_csv(STUDENT_LIST_CSV)
        for text_dir in file_list_txt:
            # print(text_dir)
            # print(1)
            dict1 = {}
            for i in range(len(df1)):
                student_info = df1.iloc[i, 0].astype(str)
                dict1[student_info] = [df1.iloc[i, 1]]
            text1 = open(text_dir, "r", encoding='UTF16')
            # print(2)
            counter=0
            for line in text1:
                # print(3)
                if line == "\n":
                    continue
                list_line = line.split()
                if len(list_line)<6:
                    continue

                # if list_line[4] == "오전":
                access_time = list_line[5]
                access_time = access_time
                # elif list_line[4] == "오후":
                    # access_time = list_line[5]
                    # hour=int(access_time[:2])
                    # hour=hour+12
                    # access_time=str(hour)+access_time[2:]
                # print(4)
                if counter==0:
                    today_check=list_line[:4]
                    counter=1

                # 이 부분이 계속 반복되서 마지막에 연속으로 채팅칠 경우에 today 변수가 날짜대신
                # 다른 텍스트로 들어가는 경우가 있음. 기능상에 문제는 없음.
                today_year=list_line[0]
                today_month=list_line[1]
                today_date=list_line[2]
                today=today_month+today_date
                today_this=list_line[:4]

                if today_this!=today_check:
                    continue

                target_list_line = set(list_line[6:])
                target_list_line = list(target_list_line)
                target_list_line.sort(key=lambda x: len(x), reverse=True)
                stu_id = ""
                for i in range(len(target_list_line)):
                    if len(target_list_line[i]) == 9 and target_list_line[i].isdigit() is True and target_list_line[i][:2] == "20":
                        for i in range(len(target_list_line)):
                            for item in dict1:
                                if dict1[item][0] in target_list_line[i]:
                                    if len(stu_id) == 0:
                                        stu_id = item
                                        break
                    else:
                        if target_list_line[i][:9].isdigit() is True and target_list_line[i][:2] == "20":
                            if len(stu_id) == 0:
                                stu_id = target_list_line[i][:9]
                                break
                        else:
                            for i in range(len(target_list_line)):
                                for item in dict1:
                                    if dict1[item][0] in target_list_line[i]:
                                        if len(stu_id) == 0:
                                            stu_id=item
                                            break
                if len(stu_id) != 0:
                    if stu_id in dict1:
                        dict1[stu_id].append(access_time)
                    elif int(stu_id) in dict1:
                        dict1[int(stu_id)].append(access_time)
            text1.close()
            df1[today] = ""
            for item in dict1:
                if len(dict1[item]) == 1:
                    df1.loc[df1["성명"] == dict1[item][0], today] = "결석"
                else:
                    first_access_time_str = dict1[item][1]
                    # print(first_access_time_str)
                    access_datetime_object = datetime.datetime.strptime(first_access_time_str, "%H:%M")
                    # print(access_datetime_object, limit_datetime_object)
                    if access_datetime_object > limit_datetime_object:
                        df1.loc[df1["성명"] == dict1[item][0], today] = "지각"
                    pass
            print("파일하나 끝남")
        df1.to_csv("result.csv", encoding="utf-8-sig")
        # ========================================= phase 2 =====================================================
        df2 = pd.read_csv("result.csv")
        t = df2.values.tolist()
        c1 = 0
        c2 = 0
        c1_list = []
        c2_list = []
        for i in range(len(t)):
            for j in range(len(t[0])):
                if t[i][j] == "지각":
                    c1 += 1
                elif t[i][j] == "결석":
                    c2 += 1
            c1_list.append(c1)
            c2_list.append(c2)
            c1 = 0
            c2 = 0
        dict2 = {"지각": c1_list, "결석": c2_list}
        df3 = pd.DataFrame(dict2)
        df4 = pd.concat([df2, df3], axis=1)
        df4.to_csv("result.csv", encoding="utf-8-sig")

        #     alert
        msg = QMessageBox()
        msg.setWindowTitle("alert")
        msg.setText("출결 기록 완료!\nresult.csv 파일을 확인해주세요")
        msg.setStandardButtons(QMessageBox.Ok)
        result = msg.exec_()

    def pushButtonClicked1(self):
        fname = QFileDialog.getOpenFileName(self)
        self.label1.setText(fname[0])

    def pushButtonClicked4(self):
        text, ok = QInputDialog.getText(self, 'Input Dialog', '예시\n오전 9시 30분 -> 09:30\n오후 3시 45분 -> 15:45')
        if ok:
            self.label4.setText(str(text))

    def pushButtonClicked2(self):
        text, ok = QInputDialog.getText(self, 'Input Dialog', '예시\n5분까지는 봐줌, 6분부터 지각 -> 5\n20분까지는 바줌, 21분부터 지각 -> 20')
        if ok:
            self.label2.setText(str(text))

if __name__ == "__main__" :
    #QApplication : 프로그램을 실행시켜주는 클래스
    app = QApplication(sys.argv)

    #WindowClass의 인스턴스 생성
    myWindow = WindowClass()

    #프로그램 화면을 보여주는 코드
    myWindow.show()

    #프로그램을 이벤트루프로 진입시키는(프로그램을 작동시키는) 코드
    app.exec_()