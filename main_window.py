# -*- coding: utf-8 -*-

import sys
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
import os
import time
import excel_util
from spider_djk import SpiderDjk
from mask_layout import MaskWidget


class MainWindow(QTabWidget):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        MainWindow.setFixedSize(self, 760, 500)

        self.cust_list = []
        self.worker_spider = ''
        self.spider = SpiderDjk()

        self.tab1 = QWidget()
        self.tab2 = QWidget()

        self.addTab(self.tab1, "Tab 1")
        self.addTab(self.tab2, "Tab 2")

        self.tab1UI()
        self.tab2UI()

        self.setWindowTitle("贷记卡系统爬取工具")
        self.setGeometry(500, 500, 500, 500)


    def tab1UI(self):
        layout = QVBoxLayout()
        sp0 = QSplitter(Qt.Horizontal)
        sp1 = QSplitter(Qt.Horizontal)
        sp2 = QSplitter(Qt.Horizontal)
        sp3 = QSplitter(Qt.Horizontal)
        sp4 = QSplitter(Qt.Horizontal)
        sp5 = QSplitter(Qt.Horizontal)
        sp6 = QSplitter(Qt.Horizontal)
        sp7 = QSplitter(Qt.Horizontal)

        lb_account = QLabel('请输入大额分期卡号：')
        self.le_account = QLineEdit()
        btn_query = QPushButton('查询')
        btn_reset = QPushButton('重置')

        btn_reset.clicked.connect(self.click_reset)
        btn_query.clicked.connect(self.click_query)

        lb_account.setFixedHeight(50)
        lb_account.setFixedHeight(50)
        lb_account.setFixedHeight(50)
        lb_account.setFixedHeight(50)
        sp0.addWidget(lb_account)
        sp0.addWidget(self.le_account)
        sp0.addWidget(btn_query)
        sp0.addWidget(btn_reset)

        lb_name = QLabel('客户名称：')
        self.le_name = QLineEdit()
        lb_certno = QLabel('    证件号码：')
        self.le_certno = QLineEdit()
        lb_name.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        lb_certno.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        lb_name.setFixedHeight(50)
        self.le_name.setFixedSize(150, 50)
        lb_certno.setFixedHeight(50)
        self.le_certno.setFixedSize(150, 50)
        sp1.addWidget(lb_name)
        sp1.addWidget(self.le_name)
        sp1.addWidget(lb_certno)
        sp1.addWidget(self.le_certno)

        lb_sum = QLabel(' 结清金额：')
        self.le_sum = QLineEdit()
        lb_exsum = QLabel('结清后溢缴额 ：')
        self.le_exsum = QLineEdit()
        lb_sum.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        lb_exsum.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        lb_sum.setFixedHeight(50)
        self.le_sum.setFixedSize(150, 50)
        lb_exsum.setFixedHeight(50)
        self.le_exsum.setFixedSize(150, 50)
        sp2.addWidget(lb_sum)
        sp2.addWidget(self.le_sum)
        sp2.addWidget(lb_exsum)
        sp2.addWidget(self.le_exsum)

        lb_bal = QLabel(' 账单余额：')
        self.le_bal = QLineEdit()
        lb_retppl = QLabel('当期还款金额 ：')
        self.le_retppl = QLineEdit()
        lb_bal.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        lb_retppl.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        lb_bal.setFixedHeight(50)
        self.le_bal.setFixedSize(150, 50)
        lb_retppl.setFixedHeight(50)
        self.le_retppl.setFixedSize(150, 50)
        sp3.addWidget(lb_bal)
        sp3.addWidget(self.le_bal)
        sp3.addWidget(lb_retppl)
        sp3.addWidget(self.le_retppl)

        lb_account_bal = QLabel('  当前账户余额：')
        self.le_account_bal = QLineEdit()
        lb_aval_exsum = QLabel('第一币种可转出溢缴额：')
        self.le_aval_exsum = QLineEdit()
        lb_account_bal.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        lb_aval_exsum.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        lb_account_bal.setFixedHeight(50)
        self.le_account_bal.setFixedSize(150, 50)
        lb_aval_exsum.setFixedHeight(50)
        self.le_aval_exsum.setFixedSize(150, 50)
        sp4.addWidget(lb_account_bal)
        sp4.addWidget(self.le_account_bal)
        sp4.addWidget(lb_aval_exsum)
        sp4.addWidget(self.le_aval_exsum)

        lb_bill_day = QLabel('         账单日：')
        self.le_bill_day = QLineEdit()
        lb_ld = QLabel('应收未收逾期还款违约金：')
        self.le_ld = QLineEdit()
        lb_bill_day.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        lb_ld.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        lb_bill_day.setFixedHeight(50)
        self.le_bill_day.setFixedSize(150, 50)
        lb_ld.setFixedHeight(50)
        self.le_ld.setFixedSize(150, 50)
        sp5.addWidget(lb_bill_day)
        sp5.addWidget(self.le_bill_day)
        sp5.addWidget(lb_ld)
        sp5.addWidget(self.le_ld)

        lb_rem_ppl = QLabel('大额分期剩余本金：')
        self.le_rem_ppl = QLineEdit()
        lb_fee = QLabel('大额分期剩余手续费/利息：')
        self.le_fee = QLineEdit()
        lb_rem_ppl.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        lb_fee.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        lb_rem_ppl.setFixedHeight(50)
        self.le_rem_ppl.setFixedSize(150, 50)
        lb_fee.setFixedHeight(50)
        self.le_fee.setFixedSize(150, 50)
        sp6.addWidget(lb_rem_ppl)
        sp6.addWidget(self.le_rem_ppl)
        sp6.addWidget(lb_fee)
        sp6.addWidget(self.le_fee)

        btn_create_customer_excel = QPushButton('生成模板')
        btn_create_customer_excel.clicked.connect(self.click_create_excel)
        btn_batch_query = QPushButton('批量查询')
        btn_batch_query.clicked.connect(self.click_batch_query)
        sp7.addWidget(btn_batch_query)

        layout.addWidget(sp0)
        layout.addWidget(sp1)
        layout.addWidget(sp2)
        layout.addWidget(sp3)
        layout.addWidget(sp4)
        layout.addWidget(sp5)
        layout.addWidget(sp6)
        layout.addWidget(sp7)
        self.setTabText(0, '查询')
        self.tab1.setLayout(layout)

    def tab2UI(self):
        layout = QVBoxLayout()

        self.setTabText(1, '设置')
        self.tab1.setLayout(layout)

    def click_query(self):
        try:
            account = self.le_account.text().strip()
            if account is None or len(account) == 0:
                self.alert_dialog('非法账号，请重新输入！！')
                return
            self.mask = MaskWidget(self)
            self.mask.show()
            l = [account]
            self.worker_spider = WorderSpider(l, self.spider)
            self.worker_spider.sig_complete.connect(self.capture_complete)
            self.worker_spider.start()
        except Exception as e:
            print(e)
            if self.mask is not None:
                self.mask.close()

    def alert_dialog(self, msg):
        try:
            QMessageBox.question(self, '提示', msg,  QMessageBox.Yes)
        except Exception as e:
            print(e)

    def click_reset(self):
        self.le_account.setText('')
        self.le_account_bal.setText('')
        self.le_aval_exsum.setText('')
        self.le_bal.setText('')
        self.le_bill_day.setText('')
        self.le_certno.setText('')
        self.le_exsum.setText('')
        self.le_fee.setText('')
        self.le_ld.setText('')
        self.le_name.setText('')
        self.le_rem_ppl.setText('')
        self.le_retppl.setText('')
        self.le_sum.setText('')

    def click_create_excel(self):
        pass

    def click_batch_query(self):
        try:
            s = QFileDialog.getOpenFileName(self, "客户导入", "/", "Excel File(*.xlsx)")

            if excel_util.is_excel_exsits(s[0]):
                self.cust_list = excel_util.get_djk_customers_from_excel(s[0])
                self.log('导入客户成功！！开始爬取客户信息\n')
                self.worker_spider = WorderSpider(self.cust_list, self.spider)
                self.worker_spider.sig_complete.connect(self.capture_complete)
                self.worker_spider.start()
            else:
                self.log('文件选取异常，请重新选择 xlsx 格式的文件！！')
        except Exception as e:
            self.infoline_log_0('程序异常 %s！！' % e)
            print(e)

    def capture_complete(self, msg):
        self.log(msg)
        if self.mask is not None:
            self.mask.close()

    def log(self, msg):
        print(msg)




class WorderSpider(QThread):
    sig_complete = pyqtSignal(str)

    def __init__(self, cust_list, spider):
        super(WorderSpider, self).__init__()
        self.cust_list = cust_list
        self.spider = spider
        print(self.cust_list)

    def run(self):
        try:
            self.spider.process()
            self.sig_complete.emit('')
        except Exception as e:
            print(e)
            self.sig_complete.emit(e)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    demo = MainWindow()
    demo.show()
    sys.exit(app.exec_())
