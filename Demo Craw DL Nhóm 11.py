#khai báo thư viện
from openpyxl import Workbook   #thư viện làm việc với excel
import requests #thư viện gửi request tới server cũng như xử lý response một cách đơn giản
import re #thư viện cung cấp các phương thức, hàm và hằng để làm việc với RegEx
from matplotlib import pyplot #Thư viện vẽ biểu đồ
from PyQt5 import QtCore, QtGui, QtWidgets #thư viện để làm việc giao diện
from PyQt5.QtWidgets import QFileDialog, QApplication, QComboBox, QVBoxLayout, QPushButton, QWidget
from PyQt5.QtGui import QPixmap
import sys
import os

#khai báo 2 list 
dataTenNuoc = []
dataDanSo = []

#mở excel
wb = Workbook()
ws = wb.active

class Ui_MainWindow(object):
    nameUrl = "/a-rap-xe-ut/"
    url = 'https://danso.org' + nameUrl
    response = requests.request("POST", url)
    txt = response.text
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(800, 598)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.Browser = QtWidgets.QPushButton(self.centralwidget)
        self.Browser.setGeometry(QtCore.QRect(512, 330, 101, 28))
        self.Browser.setObjectName("Browser")
        self.ScanButton = QtWidgets.QPushButton(self.centralwidget)
        self.ScanButton.setGeometry(QtCore.QRect(640, 330, 93, 28))
        self.ScanButton.setObjectName("ScanButton")
        self.lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit.setGeometry(QtCore.QRect(70, 320, 421, 41))
        self.lineEdit.setReadOnly(False)
        self.lineEdit.setObjectName("lineEdit")
        self.lineEdit_2 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_2.setGeometry(QtCore.QRect(20, 20, 231, 71))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.lineEdit_2.setFont(font)
        self.lineEdit_2.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignTop)
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(20, 40, 231, 16))
        self.label_2.setObjectName("label_2")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(20, 60, 231, 20))
        self.label.setObjectName("label")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(210, 160, 401, 71))
        font = QtGui.QFont()
        font.setPointSize(18)
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_3")
        self.buttonShowChart = QtWidgets.QPushButton(self.centralwidget)
        self.buttonShowChart.setGeometry(QtCore.QRect(640, 380, 93, 28))
        self.buttonShowChart.setObjectName("buttonShowChart")
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setGeometry(QtCore.QRect(680, 520, 111, 16))
        self.label_4.setObjectName("label_4")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 800, 26))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        #----------------------------------------
        #Nhấn nút chọn đường dẫn->gọi hàm getLink
        self.Browser.clicked.connect(self.getLink)
        #Nhấn nút Craw->gọi hàm clickCraw
        self.ScanButton.clicked.connect(self.clickCraw)
        #Nhấn nút Xem biểu đồ->gọi hàm showBieuDo
        self.buttonShowChart.clicked.connect(self.showBieuDo)
        #---------------
        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "CrawData"))
        self.Browser.setText(_translate("MainWindow", "Chọn đường dẫn"))
        self.ScanButton.setText(_translate("MainWindow", "Craw"))
        self.lineEdit_2.setText(_translate("MainWindow", "Hướng dẫn sử dụng:   "))
        self.label_2.setText(_translate("MainWindow", " 1.Chọn đường dẫn lưu file"))
        self.label.setText(_translate("MainWindow", " 2.Ấn nút Craw"))
        self.label_3.setText(_translate("MainWindow", "DEMO CRAW DATA DÂN SỐ"))
        self.buttonShowChart.setText(_translate("MainWindow", "Xem biểu đồ"))
        self.label_4.setText(_translate("MainWindow", "Made by Group 11"))

    #Hàm lấy đường dẫn lưu
    def getLink(self):
        filename = QFileDialog.getSaveFileName(None, "craw" , "craw data.xlsx", 'Excel File(.xlsx)' , 'Excel File(.xlsx)')
        self.lineEdit.setText(filename[0])

    #hàm lấy data của 1 nước
    def layData(self):
        #Lấy tên nước
        tenNuoc = r"(?<=của <strong>).*?(?=</strong>)"
        ketqua1 = re.findall(tenNuoc, self.txt)
        print("-----" + str(ketqua1) + "-----")
        #Lấy dân số 
        dsCurrent  = r"(?<=là <strong>).*?(?=</strong> người vào ngày)"
        ketqua2 = re.findall(dsCurrent, self.txt)
        print("Dân số hiện tại: " + str(ketqua2))
        #Lấy % dân so với thế giới
        chiem = r"(?<=hiện chiếm <strong>).*?(?=</strong>)"
        ketqua3 = re.findall(chiem, self.txt)
        print("Chiếm: "+ str(ketqua3) + " dân số thế giới")
        #Lấy xếp hạng
        xepHang = r"(?<=đứng thứ <strong>).*?(?=</strong> trên)"
        ketqua4 = re.findall(xepHang, self.txt)
        print("Đứng thứ: "+ str(ketqua4) + " trên thế giới")
        #Lấy mật độ dân số
        matDo = r"(?<=" + str(ketqua1[0]) + " là <strong>).*?(?=</strong> người/km)"
        ketqua5 = re.findall(matDo, self.txt)
        print("Mật độ dân số: "+ str(ketqua5) + " người/km vuông")
        #Lấy tổng diện tích đất
        tongDienTichDat = r"(?<=>đất</strong> là <strong>).*?(?=km<sup>2</sup>)"
        ketqua6= re.findall(tongDienTichDat, self.txt)
        print("Tong dien tich dat: "+ str(ketqua6) + " km^2")
        #Lấy số dân sống ở thành thị
        danSongThanhThi = r"(?<=></strong>.</li><li><strong>).*?(?=</strong>)"
        ketqua7= re.findall(danSongThanhThi, self.txt)
        print("Dan song thanh thi: "+ str(ketqua7))
        #Lấy tuổi trung bình
        doTuoiTrungBinh = f"(?<=ở " + str(ketqua1[0]) + " là <strong>).*?(?=</strong> tuổi)"
        ketqua8 = re.findall(doTuoiTrungBinh, self.txt)
        print("Độ tuổi trung bình: "+ str(ketqua8))  

        #Convert qua số
        if(ketqua2[0] != '0'):   
            #print("khac")
            chuoimoi = ketqua2[0].replace('.', '')
            s_ = int(float(chuoimoi))
            #print(s_)
            dataDanSo.append(s_)
        else:   
            dataDanSo.append(0)
            #print("giong")

        #Thêm vào list
        dataTenNuoc.append(str(ketqua1[0]))
        #Thêm data(tên nước,dân số,..) vào excel
        data1 = [[str(ketqua1[0]), str(ketqua2[0]), str(ketqua3[0]), str(ketqua4[0]), str(ketqua5[0]), str(ketqua6[0]), str(ketqua7[0]), str(ketqua8[0])]]
        for data in data1:
            ws.append(data)
        wb.save(self.lineEdit.text())

    #hàm vẽ biểu đồ
    def showBieuDo(self):
        pyplot.ticklabel_format(axis="y", style='plain')
        pyplot.title("Biểu đồ dân số")
        pyplot.bar(dataTenNuoc, dataDanSo)
        pyplot.xlabel("Quốc gia")
        pyplot.ylabel("Dân Số(người)")
        pyplot.grid(True)
        pyplot.show()     

    #hàm gọi khi nhấn vào nút craw
    def clickCraw(self):
        #hiển thị hộp thông báo người dùng đã nhấn nút craw
        msg = QtWidgets.QMessageBox()
        msg.setInformativeText("Đã xác nhận craw,vui lòng nhấn OK để bắt đầu")
        msg.exec()    
        #thêm câu dưới vào dòng đầu trong excel
        if(self.lineEdit.text() != None):
            testData = [["Tên nước","Dân số hiện tại","Chiếm % so với thế giới", "Đứng thứ", "Mật độ dân số", "Tổng diện tích đất", "Dân sống thành thị", "Độ tuổi trung bình"]]
            for data in testData:
                ws.append(data)
            wb.save(self.lineEdit.text())
        #lấy option value
        tenVaValue = f'(?<=</option><option value=").*?(?=</option>)'
        ketqua9 = re.findall(tenVaValue, self.txt)
        #tìm kí tự ">" để lấy mỗi kí tự trong dấu /
        for i in range(len(ketqua9)):
            name_1 = ketqua9[i]
            pos_ = name_1.find('">')
            chuoimoi = ""
            j = 0;
            for k in range(len(name_1)):
                if(j < pos_):
                    chuoimoi += name_1[j]
                    j += 1    
            nameUrl1 = chuoimoi
            self.nameUrl = nameUrl1
            self.url = 'https://danso.org' + nameUrl1
            self.response = requests.request("POST", self.url)
            self.txt = self.response.text
            self.layData()
            #hiển thị hộp thông báo khi craw xong
            if (i >= len(ketqua9)):
                msg = QtWidgets.QMessageBox()
                msg.setInformativeText("Đã xong")
                msg.exec()    

if _name_ == "_main_":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())