import sys
import json
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QLabel, QLineEdit, QComboBox, 
                            QPushButton, QMessageBox, QDialog, QListWidget)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os

# 定義樣式
STYLE_SHEET = """
QMainWindow, QDialog {
    background-color: #f0f0f0;
}
QLabel {
    font-size: 12px;
    color: #333333;
}
QLineEdit {
    padding: 5px;
    border: 1px solid #cccccc;
    border-radius: 3px;
    background-color: white;
    min-height: 25px;
}
QComboBox {
    padding: 5px;
    border: 1px solid #cccccc;
    border-radius: 3px;
    background-color: white;
    min-height: 25px;
}
QPushButton {
    background-color: #4a90e2;
    color: white;
    border: none;
    padding: 8px 15px;
    border-radius: 3px;
    min-height: 30px;
}
QPushButton:hover {
    background-color: #357abd;
}
QPushButton:pressed {
    background-color: #2a5885;
}
QListWidget {
    border: 1px solid #cccccc;
    border-radius: 3px;
    background-color: white;
}
QListWidget::item {
    padding: 5px;
}
QListWidget::item:selected {
    background-color: #4a90e2;
    color: white;
}
"""
#color . 

class MemberManageDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle('會員資料管理')
        self.setGeometry(350, 350, 400, 500)
        self.setStyleSheet(STYLE_SHEET)
        
        layout = QVBoxLayout()
        layout.setSpacing(10)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # 會員列表
        list_label = QLabel('已保存的會員:')
        list_label.setFont(QFont('Arial', 12, QFont.Bold))
        layout.addWidget(list_label)
        
        self.members_list = QListWidget()
        self.update_members_list()
        layout.addWidget(self.members_list)
        
        # 刪除按鈕
        self.delete_button = QPushButton('刪除選中的會員')
        self.delete_button.setStyleSheet("""
            QPushButton {
                background-color: #e74c3c;
            }
            QPushButton:hover {
                background-color: #c0392b;
            }
        """)
        self.delete_button.clicked.connect(self.delete_member)
        layout.addWidget(self.delete_button)
        
        # 分隔線
        separator = QLabel('新增會員')
        separator.setAlignment(Qt.AlignCenter)
        separator.setFont(QFont('Arial', 12, QFont.Bold))
        separator.setStyleSheet("QLabel { background-color: #e0e0e0; padding: 8px; border-radius: 3px; }")
        layout.addWidget(separator)
        
        # 編碼輸入/手動
        member_layout = QHBoxLayout()
        member_label = QLabel('會員代號:')
        self.member_input = QLineEdit()
        member_layout.addWidget(member_label)
        member_layout.addWidget(self.member_input)
        layout.addLayout(member_layout)
        
        # 會員備註輸入
        note_layout = QHBoxLayout()
        note_label = QLabel('備註說明:')
        self.note_input = QLineEdit()
        note_layout.addWidget(note_label)
        note_layout.addWidget(self.note_input)
        layout.addLayout(note_layout)
        
        # 保存按鈕
        self.save_button = QPushButton('保存新會員')
        self.save_button.setStyleSheet("""
            QPushButton {
                background-color: #2ecc71;
            }
            QPushButton:hover {
                background-color: #27ae60;
            }
        """) #顏色
        self.save_button.clicked.connect(self.save_member)
        layout.addWidget(self.save_button)
        
        self.setLayout(layout)
    
    def update_members_list(self):
        self.members_list.clear()
        for member_id, data in self.parent.members_data.items():
            display_text = f"{member_id}({data['note']})"
            self.members_list.addItem(display_text)
    
    def delete_member(self):
        current_item = self.members_list.currentItem()
        if not current_item:
            QMessageBox.warning(self, '警告', '請選擇要刪除的會員')
            return
            
        member_id = current_item.text().split('(')[0]
        reply = QMessageBox.question(self, '確認刪除', 
                                   f'確定要刪除會員 {current_item.text()} 嗎？',
                                   QMessageBox.Yes | QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            del self.parent.members_data[member_id]
            self.parent.save_members_data()
            self.update_members_list()
            self.parent.update_member_combo()
            QMessageBox.information(self, '成功', '會員已刪除')
    
    def save_member(self):
        member_id = self.member_input.text().strip()
        member_note = self.note_input.text().strip()
        
        if not member_id:
            QMessageBox.warning(self, '警告', '請輸入會員編號')
            return
            
        if member_id in self.parent.members_data:
            reply = QMessageBox.question(self, '確認覆蓋', 
                                       f'會員 {member_id} 已存在，是否覆蓋？',
                                       QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.No:
                return #編號ID
        
        self.parent.members_data[member_id] = {
            'note': member_note
        }
        
        self.parent.save_members_data()
        self.parent.update_member_combo()
        self.update_members_list()
        
        QMessageBox.information(self, '成功', '會員資料已保存')
        self.member_input.clear()
        self.note_input.clear() #json.掛上去的

class MapleStoryForm(QMainWindow):
    def __init__(self):
        super().__init__()
        self.members_data = {}
        self.load_members_data()
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle('MapleStory M Coupon Redeemer')
        self.setGeometry(300, 300, 400, 280)  # 增加窗口高度以容納署名
        self.setStyleSheet(STYLE_SHEET)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setSpacing(10)
        layout.setContentsMargins(20, 20, 20, 20)

        # 會員選擇
        member_select_layout = QHBoxLayout()
        member_select_label = QLabel('選擇會員:')
        self.member_select_combo = QComboBox()
        self.update_member_combo()
        self.manage_member_button = QPushButton('管理會員')
        self.manage_member_button.setStyleSheet("""
            QPushButton {
                background-color: #9b59b6;
            }
            QPushButton:hover {
                background-color: #8e44ad;
            }
        """)
        self.manage_member_button.clicked.connect(self.show_member_dialog)
        member_select_layout.addWidget(member_select_label)
        member_select_layout.addWidget(self.member_select_combo)
        member_select_layout.addWidget(self.manage_member_button)
        layout.addLayout(member_select_layout)

        # 服務器選擇
        server_layout = QHBoxLayout()
        server_label = QLabel('Server:')
        self.server_combo = QComboBox()
        self.server_combo.addItems(['', 'Asia1', 'Asia2'])
        server_layout.addWidget(server_label)
        server_layout.addWidget(self.server_combo)
        layout.addLayout(server_layout)

        # 序號輸入
        serial_layout = QHBoxLayout()
        serial_label = QLabel('請輸入序號:')
        self.serial_input = QLineEdit()
        serial_layout.addWidget(serial_label)
        serial_layout.addWidget(self.serial_input)
        layout.addLayout(serial_layout)

        # 提交按鈕
        self.submit_button = QPushButton('確認兌換')
        self.submit_button.clicked.connect(self.submit_form)
        layout.addWidget(self.submit_button)

        # 狀態顯示
        self.status_label = QLabel('')
        self.status_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.status_label)

        # 底部logo
        signature_label = QLabel('By Asia1 露娜 LycoRec 墨燼染月.')
        signature_label.setAlignment(Qt.AlignRight)
        signature_label.setStyleSheet("""
            QLabel {
                color: #666666;
                font-style: italic;
                margin-top: 10px;
            }
        """) #顏色
        layout.addWidget(signature_label)

    def show_member_dialog(self):
        dialog = MemberManageDialog(self)
        dialog.exec_()

    def load_members_data(self):
        try:
            if os.path.exists('members_data.json'):
                with open('members_data.json', 'r', encoding='utf-8') as f:
                    self.members_data = json.load(f)
        except Exception as e:
            print(f"載入會員資料時發生錯誤: {e}")
            self.members_data = {}

    def save_members_data(self):
        try:
            with open('members_data.json', 'w', encoding='utf-8') as f:
                json.dump(self.members_data, f, ensure_ascii=False, indent=4)
        except Exception as e:
            QMessageBox.warning(self, '錯誤', f'儲存會員資料時發生錯誤: {e}')

    def update_member_combo(self):
        self.member_select_combo.clear()
        self.member_select_combo.addItem('')
        for member_id, data in self.members_data.items():
            display_text = f"{member_id}({data['note']})"
            self.member_select_combo.addItem(display_text)

    def submit_form(self):
        try:
            server = self.server_combo.currentText()
            selected_text = self.member_select_combo.currentText()
            member_id = selected_text.split('(')[0] if selected_text else ''
            serial_code = self.serial_input.text()

            if not all([server, member_id, serial_code]):
                self.status_label.setText('請填寫所有必要資訊')
                return

            self.status_label.setText('請稍等...')
            self.automate_web(server, member_id, serial_code)
            
        except Exception as e:
            self.status_label.setText(f'請稍後: {str(e)}')

    def automate_web(self, server, member_id, serial_code):
        options = webdriver.ChromeOptions()
        options.add_argument('--disable-gpu')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-extensions')
        options.add_argument('--disable-logging')
        options.add_argument('--disable-notifications')
        options.add_argument('--disable-default-apps')
        driver = webdriver.Chrome(options=options)
        
        try:
            driver.get('https://mcoupon.nexon.com/maplestorym_global?lang=zh-TW')
            wait = WebDriverWait(driver, 5)
            
            server_dropdown = wait.until(
                EC.element_to_be_clickable((By.ID, "eRedeemRegion"))
            )
            Select(server_dropdown).select_by_value(server)
            time.sleep(0.3)
            
            member_field = wait.until(
                EC.presence_of_element_located((By.ID, "eRedeemNpaCode"))
            )
            driver.execute_script("arguments[0].value = arguments[1];", member_field, member_id)
            time.sleep(0.3)
            
            serial_field = wait.until(
                EC.presence_of_element_located((By.ID, "eRedeemCoupon"))
            )
            driver.execute_script("arguments[0].value = arguments[1];", serial_field, serial_code)
            time.sleep(0.3)

            try:
                confirm_button = wait.until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, "button.btn_confirm.e-characters-with-npacode[data-message='redeem']")
                    )
                )
                driver.execute_script("arguments[0].click();", confirm_button)
                time.sleep(2)
                self.status_label.setText('開啟兌換成功...')
            except Exception as e:
                self.status_label.setText(f'失敗: {str(e)}')
                return
            
        except Exception as e:
            self.status_label.setText(f'過程出錯: {str(e)}')
        # finally:
        #     driver.quit()  #times  好像用不到(?) else? 
if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    ex = MapleStoryForm()
    ex.show()
    sys.exit(app.exec_())