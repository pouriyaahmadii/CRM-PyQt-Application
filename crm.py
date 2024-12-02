import sys
import sqlite3
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
                             QLineEdit, QPushButton, QTableWidget, QTableWidgetItem, QMessageBox, 
                             QFileDialog, QTextEdit, QComboBox, QHeaderView, QDialog, QCalendarWidget)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QPixmap, QFont
import jdatetime
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import pandas as pd
from PyQt5.QtWidgets import QDateEdit
from PyQt5.QtCore import QDate

class CRM(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("مزون سارن")
        self.setGeometry(100, 100, 1000, 700)
        self.setLayoutDirection(Qt.RightToLeft)
        
        self.setStyleSheet("""
            QWidget {
                font-family: Arial;
                font-size: 14px;
                font-weight: bold;
            }
            QLabel {
                font-size: 16px;
            }
            QPushButton {
                font-size: 16px;
            }
        """)
        
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)
        
        self.create_widgets()
        self.create_database()
        self.search_data()

    def create_widgets(self):
        form_layout = QHBoxLayout()
        
        left_layout = QVBoxLayout()
        right_layout = QVBoxLayout()
        
        self.name_entry = QLineEdit()
        left_layout.addWidget(QLabel("نام:"))
        left_layout.addWidget(self.name_entry)
        
        self.surname_entry = QLineEdit()
        left_layout.addWidget(QLabel("نام خانوادگی:"))
        left_layout.addWidget(self.surname_entry)
        
        self.contract_date_entry = QLineEdit()
        self.contract_date_entry.setPlaceholderText("YYYY/MM/DD")
        self.contract_date_entry.textChanged.connect(self.format_date)
        left_layout.addWidget(QLabel("تاریخ عقد قرارداد:"))
        left_layout.addWidget(self.contract_date_entry)
        
        self.phone_entry = QLineEdit()
        left_layout.addWidget(QLabel("شماره موبایل:"))
        left_layout.addWidget(self.phone_entry)
        
        self.address_entry = QLineEdit()
        right_layout.addWidget(QLabel("آدرس:"))
        right_layout.addWidget(self.address_entry)
        
        self.description_entry = QTextEdit()
        self.description_entry.setMaximumHeight(100)
        right_layout.addWidget(QLabel("توضیحات:"))
        right_layout.addWidget(self.description_entry)
        
        self.total_price_entry = QLineEdit()
        self.total_price_entry.textChanged.connect(self.format_price)
        right_layout.addWidget(QLabel("قیمت کل:"))
        right_layout.addWidget(self.total_price_entry)
        
        self.final_price_entry = QLineEdit()
        self.final_price_entry.textChanged.connect(self.format_price)
        right_layout.addWidget(QLabel("قیمت نهایی:"))
        right_layout.addWidget(self.final_price_entry)
        
        self.date_entry = QLineEdit()
        self.date_entry.setPlaceholderText("YYYY/MM/DD")
        self.date_entry.textChanged.connect(self.format_date)
        right_layout.addWidget(QLabel("تاریخ ثبت:"))
        right_layout.addWidget(self.date_entry)

        self.product_entry = QLineEdit()
        right_layout.addWidget(QLabel("کالا:"))
        right_layout.addWidget(self.product_entry)

        self.product_type = QLineEdit()
        right_layout.addWidget(QLabel("نوع کالا:"))
        right_layout.addWidget(self.product_type)

        self.return_date_entry = QLineEdit()
        self.return_date_entry.setPlaceholderText("YYYY/MM/DD")
        self.return_date_entry.textChanged.connect(self.format_date)
        right_layout.addWidget(QLabel("تاریخ برگشت کالا:"))
        right_layout.addWidget(self.return_date_entry)
        
        form_layout.addLayout(left_layout)
        form_layout.addLayout(right_layout)
        
        self.layout.addLayout(form_layout)
        
        button_layout = QHBoxLayout()
        self.save_button = QPushButton("ذخیره")
        self.save_button.clicked.connect(self.save_data)
        button_layout.addWidget(self.save_button)
    
        self.pdf_button = QPushButton("خروجی PDF")
        self.pdf_button.clicked.connect(self.export_pdf)
        button_layout.addWidget(self.pdf_button)
    
        self.excel_button = QPushButton("خروجی Excel")
        self.excel_button.clicked.connect(self.export_excel)
        button_layout.addWidget(self.excel_button)
    
        self.layout.addLayout(button_layout)
        
        search_layout = QHBoxLayout()
        self.search_entry = QLineEdit()
        self.search_entry.setPlaceholderText("جستجو...")
        search_layout.addWidget(self.search_entry)
        
        self.filter_combo = QComboBox()
        self.filter_combo.addItems(["همه", "روزانه", "هفتگی", "ماهانه", "سالانه", "بازه زمانی"])
        search_layout.addWidget(self.filter_combo)
        
        self.start_date = QLineEdit()
        self.start_date.setPlaceholderText("YYYY/MM/DD")
        self.start_date.textChanged.connect(self.format_date)
        self.start_date.hide()
        search_layout.addWidget(self.start_date)

        self.end_date = QLineEdit()
        self.end_date.setPlaceholderText("YYYY/MM/DD")
        self.end_date.textChanged.connect(self.format_date)
        self.end_date.hide()
        search_layout.addWidget(self.end_date)
        
        search_button = QPushButton("جستجو")
        search_button.clicked.connect(self.search_data)
        search_layout.addWidget(search_button)
        
        self.layout.addLayout(search_layout)
        
        self.filter_combo.currentTextChanged.connect(self.toggle_date_range)
        
        self.table = QTableWidget()
        self.table.setColumnCount(15)
        self.table.setHorizontalHeaderLabels(["شناسه", "نام", "نام خانوادگی", "تاریخ عقد قرارداد", "شماره موبایل", "آدرس", "تاریخ", "توضیحات", "قیمت کل", "قیمت نهایی", "کالا", "نوع کالا", "تاریخ برگشت کالا", "ویرایش", "حذف"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.layout.addWidget(self.table)

    def format_date(self):
        sender = self.sender()
        cursor_pos = sender.cursorPosition()
        current_text = sender.text().replace('/', '')
        
        formatted_text = ''
        for i, char in enumerate(current_text):
            if i in [4, 6]:
                formatted_text += '/'
            formatted_text += char
        
        sender.setText(formatted_text)
        
        new_cursor_pos = cursor_pos + formatted_text[:cursor_pos].count('/') - current_text[:cursor_pos].count('/')
        sender.setCursorPosition(new_cursor_pos)

    def format_price(self):
        sender = self.sender()
        cursor_pos = sender.cursorPosition()
        current_text = sender.text().replace(',', '')
        
        try:
            if current_text:
                formatted_text = '{:,}'.format(int(current_text))
                sender.setText(formatted_text)
                
                new_cursor_pos = cursor_pos + formatted_text.count(',', 0, cursor_pos)
                sender.setCursorPosition(new_cursor_pos)
        except ValueError:
            pass

    def create_database(self):
        conn = sqlite3.connect("crm9.db")
        cursor = conn.cursor()
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS customers (
            id INTEGER PRIMARY KEY,
            name TEXT,
            surname TEXT,
            contract_date TEXT,
            phone TEXT,
            address TEXT,
            date TEXT,
            description TEXT,
            total_price REAL,
            final_price REAL,
            product TEXT,
            product_type TEXT,
            return_date TEXT
        )
        """)
        conn.commit()
        conn.close()

    def save_data(self):
        try:
            name = self.name_entry.text()
            surname = self.surname_entry.text()
            contract_date = self.contract_date_entry.text()
            phone = self.phone_entry.text()
            address = self.address_entry.text()
            description = self.description_entry.toPlainText()
            total_price = float(self.total_price_entry.text().replace(',', '')) if self.total_price_entry.text() else 0
            final_price = float(self.final_price_entry.text().replace(',', '')) if self.final_price_entry.text() else 0
            date = self.date_entry.text()
            product = self.product_entry.text()
            product_type = self.product_type.text()
            return_date = self.return_date_entry.text()

            conn = sqlite3.connect("crm9.db")
            cursor = conn.cursor()
            cursor.execute("INSERT INTO customers (name, surname, contract_date, phone, address, date, description, total_price, final_price, product, product_type, return_date) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                           (name, surname, contract_date, phone, address, date, description, total_price, final_price, product, return_date, product_type))
            conn.commit()
            conn.close()

            self.clear_entries()
            self.search_data()
            QMessageBox.information(self, "موفقیت", "اطلاعات با موفقیت ذخیره شد!")
        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا در ذخیره اطلاعات: {str(e)}")

    def clear_entries(self):
        self.name_entry.clear()
        self.surname_entry.clear()
        self.contract_date_entry.clear()
        self.phone_entry.clear()
        self.address_entry.clear()
        self.description_entry.clear()
        self.total_price_entry.clear()
        self.final_price_entry.clear()
        self.date_entry.clear()
        self.product_entry.clear()
        self.product_type.clear()
        self.return_date_entry.clear()

    def search_data(self):
        search_text = self.search_entry.text()
        filter_option = self.filter_combo.currentText()
        
        conn = sqlite3.connect("crm9.db")
        cursor = conn.cursor()
        
        query = "SELECT * FROM customers WHERE 1=1"
        params = []
        
        if search_text:
            query += " AND (name LIKE ? OR surname LIKE ? OR phone LIKE ?)"
            params.extend([f"%{search_text}%"] * 3)
        
        if filter_option == "روزانه":
            today = jdatetime.date.today().strftime("%Y/%m/%d")
            query += " AND date = ?"
            params.append(today)
        elif filter_option == "هفتگی":
            today = jdatetime.date.today()
            start_of_week = today - jdatetime.timedelta(days=today.weekday())
            end_of_week = start_of_week + jdatetime.timedelta(days=6)
            query += " AND date BETWEEN ? AND ?"
            params.extend([start_of_week.strftime("%Y/%m/%d"), end_of_week.strftime("%Y/%m/%d")])
        elif filter_option == "ماهانه":
            today = jdatetime.date.today()
            start_of_month = today.replace(day=1)
            end_of_month = start_of_month.replace(day=29) + jdatetime.timedelta(days=4)
            end_of_month = end_of_month - jdatetime.timedelta(days=end_of_month.day)
            query += " AND date BETWEEN ? AND ?"
            params.extend([start_of_month.strftime("%Y/%m/%d"), end_of_month.strftime("%Y/%m/%d")])
        elif filter_option == "سالانه":
            today = jdatetime.date.today()
            start_of_year = today.replace(month=1, day=1)
            end_of_year = today.replace(month=12, day=29)
            query += " AND date BETWEEN ? AND ?"
            params.extend([start_of_year.strftime("%Y/%m/%d"), end_of_year.strftime("%Y/%m/%d")])
        elif filter_option == "بازه زمانی":
            start_date = self.start_date.text()
            end_date = self.end_date.text()
            if start_date and end_date:
                query += " AND date BETWEEN ? AND ?"
                params.extend([start_date, end_date])
        
        cursor.execute(query, params)
        results = cursor.fetchall()
        
        self.table.setRowCount(0)
        for row_number, row_data in enumerate(results):
            self.table.insertRow(row_number)
            for column_number, data in enumerate(row_data):
                if column_number in [8, 9]:  # قیمت کل و قیمت نهایی
                    formatted_data = '{:,}'.format(int(data))
                    self.table.setItem(row_number, column_number, QTableWidgetItem(formatted_data))
                elif column_number == 11:  # نوع کالا
                    self.table.setItem(row_number, column_number, QTableWidgetItem(str(row_data[12])))
                elif column_number == 12:  # تاریخ برگشت کالا
                    self.table.setItem(row_number, column_number, QTableWidgetItem(str(row_data[11])))
                else:
                    self.table.setItem(row_number, column_number, QTableWidgetItem(str(data)))
            
            edit_button = QPushButton("ویرایش")
            edit_button.clicked.connect(lambda _, row=row_number: self.edit_row(row))
            self.table.setCellWidget(row_number, 13, edit_button)
            
            delete_button = QPushButton("حذف")
            delete_button.clicked.connect(lambda _, row=row_number: self.delete_row(row))
            self.table.setCellWidget(row_number, 14, delete_button)
        
        conn.close()

    def edit_row(self, row):
        item_id = self.table.item(row, 0).text()
        
        conn = sqlite3.connect("crm9.db")
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM customers WHERE id=?", (item_id,))
        customer = cursor.fetchone()
        conn.close()
        
        if customer:
            self.name_entry.setText(customer[1])
            self.surname_entry.setText(customer[2])
            self.contract_date_entry.setText(customer[3])
            self.phone_entry.setText(customer[4])
            self.address_entry.setText(customer[5])
            self.date_entry.setText(customer[6])
            self.description_entry.setText(customer[7])
            self.total_price_entry.setText('{:,}'.format(int(customer[8])))
            self.final_price_entry.setText('{:,}'.format(int(customer[9])))
            self.product_entry.setText(customer[10])
            self.product_type.setText(customer[12])
            self.return_date_entry.setText(customer[11])
            
            self.save_button.clicked.disconnect()
            self.save_button.clicked.connect(lambda: self.update_data(item_id))
            self.save_button.setText("بروزرسانی")

    def update_data(self, item_id):
        try:
            name = self.name_entry.text()
            surname = self.surname_entry.text()
            contract_date = self.contract_date_entry.text()
            phone = self.phone_entry.text()
            address = self.address_entry.text()
            description = self.description_entry.toPlainText()
            total_price = float(self.total_price_entry.text().replace(',', '')) if self.total_price_entry.text() else 0
            final_price = float(self.final_price_entry.text().replace(',', '')) if self.final_price_entry.text() else 0
            date = self.date_entry.text()
            product = self.product_entry.text()
            product_type = self.product_type.text()
            return_date = self.return_date_entry.text()

            conn = sqlite3.connect("crm9.db")
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE customers 
                SET name=?, surname=?, contract_date=?, phone=?, address=?, date=?, description=?, 
                    total_price=?, final_price=?, product=?, product_type=?, return_date=?
                WHERE id=?
            """, (name, surname, contract_date, phone, address, date, description, 
                  total_price, final_price, product, return_date, product_type, item_id))
            conn.commit()
            conn.close()

            self.clear_entries()
            self.search_data()
            self.save_button.clicked.disconnect()
            self.save_button.clicked.connect(self.save_data)
            self.save_button.setText("ذخیره")
            QMessageBox.information(self, "موفقیت", "اطلاعات با موفقیت بروزرسانی شد!")
        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا در بروزرسانی اطلاعات: {str(e)}")

    def delete_row(self, row):
        item_id = self.table.item(row, 0).text()
        reply = QMessageBox.question(self, 'حذف', 'آیا مطمئن هستید که می‌خواهید این مورد را حذف کنید؟', 
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            conn = sqlite3.connect("crm9.db")
            cursor = conn.cursor()
            cursor.execute("DELETE FROM customers WHERE id=?", (item_id,))
            conn.commit()
            conn.close()
            self.search_data()

    def toggle_date_range(self, text):
        if text == "بازه زمانی":
            self.start_date.show()
            self.end_date.show()
        else:
            self.start_date.hide()
            self.end_date.hide()

    def export_pdf(self):
        file_name, _ = QFileDialog.getSaveFileName(self, "ذخیره فایل PDF", "", "PDF Files (*.pdf)")
        if file_name:
            pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))
            c = canvas.Canvas(file_name, pagesize=letter)
            c.setFont("Arial", 14)

            y = 750
            for row in range(self.table.rowCount()):
                if y < 50:
                    c.showPage()
                    y = 750
                for col in range(self.table.columnCount() - 2):  # Exclude edit and delete columns
                    item = self.table.item(row, col)
                    if item is not None:
                        c.drawRightString(580 - col * 45, y, str(item.text()))
                y -= 20
            c.save()

    def export_excel(self):
        file_name, _ = QFileDialog.getSaveFileName(self, "ذخیره فایل Excel", "", "Excel Files (*.xlsx)")
        if file_name:
            data = []
            for row in range(self.table.rowCount()):
                row_data = []
                for col in range(self.table.columnCount() - 2):  # Exclude edit and delete columns
                    item = self.table.item(row, col)
                    if item is not None:
                        row_data.append(str(item.text()))
                    else:
                        row_data.append("")
                data.append(row_data)
            
            df = pd.DataFrame(data, columns=[self.table.horizontalHeaderItem(i).text() for i in range(self.table.columnCount() - 2)])
            df.to_excel(file_name, index=False)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = CRM()
    window.show()
    sys.exit(app.exec_())