import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QLineEdit, QMessageBox, QDialog, QFormLayout,
    QDialogButtonBox, QTextEdit, QProgressBar
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont

from parsed_info import process_companies
from constants import  EXCEL_OUTPUT

class TokenDialog(QDialog):
    def __init__(self, token_num, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Изменить TOKEN{token_num}")
        self.setFixedSize(400, 150)
        
        layout = QFormLayout()
        self.line_edit = QLineEdit()
        self.line_edit.setEchoMode(QLineEdit.Password)
        layout.addRow(f"Новый TOKEN{token_num}:", self.line_edit)
        
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addRow(buttons)
        self.setLayout(layout)

class DotParsingDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Парсинг DOT со страницы")
        self.setFixedSize(600, 250)
        
        layout = QVBoxLayout()
        label = QLabel("Введите URL страницы для парсинга DOT:")
        label.setFont(QFont("Arial", 12))
        layout.addWidget(label)
        
        self.url_input = QLineEdit()
        self.url_input.setFont(QFont("Arial", 11))
        layout.addWidget(self.url_input)
        
        self.status = QLabel("")
        self.status.setStyleSheet("color: blue;")
        layout.addWidget(self.status)
        
        btn = QPushButton("🚀 Начать парсинг DOT")
        btn.setStyleSheet("background-color: #27ae60; color: white; padding: 10px;")
        btn.clicked.connect(self.start_parsing)
        layout.addWidget(btn)
        
        self.setLayout(layout)

    def start_parsing(self):
        url = self.url_input.text().strip()
        if not url:
            QMessageBox.warning(self, "Ошибка", "Введите URL!")
            return
        
        self.status.setText("Парсинг запущен... Ожидайте...")
        self.status.setStyleSheet("color: orange;")
        QApplication.processEvents()
        
        try:
            # Здесь  функция парсинга DOT
            from Dot_pars import extract_company_dots
            result = extract_company_dots(url.strip())
            self.status.setText(result)
            self.status.setStyleSheet("color: green;")
            QMessageBox.information(self, "Успех", result)
        except Exception as e:
            self.status.setText(f"Ошибка: {str(e)}")
            self.status.setStyleSheet("color: red;")
            QMessageBox.critical(self, "Ошибка", str(e))

class ParsingThread(QThread):
    finished = pyqtSignal(str)
    error = pyqtSignal(str)
    progress = pyqtSignal(int)

    def __init__(self, excel_path):
        super().__init__()
        self.excel_path = excel_path

    def run(self):
        try:
            result = process_companies(from_excel=self.excel_path)
            self.finished.emit(result)
        except Exception as e:
            self.error.emit(str(e))

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Парсер TruckStop")
        self.setGeometry(200, 100, 700, 800)
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Заголовок
        title = QLabel("Парсер данных по DOT")
        title.setFont(QFont("Arial", 18, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
        
        # Выбор файла
        self.file_btn = QPushButton("📂 Выбрать Excel с DOT")
        self.file_btn.setFont(QFont("Arial", 12))
        self.file_btn.clicked.connect(self.select_excel)
        layout.addWidget(self.file_btn)
        
        self.file_label = QLabel("Файл не выбран")
        self.file_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.file_label)
        

        
        # Кнопка запуска
        self.parse_btn = QPushButton("🚀 Запустить парсинг")
        self.parse_btn.setFont(QFont("Arial", 14, QFont.Bold))
        self.parse_btn.setStyleSheet("background-color: #28a745; color: white; padding: 15px;")
        self.parse_btn.clicked.connect(self.run_parsing)
        layout.addWidget(self.parse_btn)
        
        # Лог
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont("Consolas", 10))
        layout.addWidget(QLabel("Лог обработки:"))
        layout.addWidget(self.log_text)
        
        layout.addStretch(1)
        
        
        
        dot_btn = QPushButton("🌐 Парсинг DOT со страницы")
        dot_btn.setStyleSheet("background-color: #3498db; color: white; padding: 12px;")
        dot_btn.clicked.connect(self.open_dot_parsing)
        layout.addWidget(dot_btn)
        
        cookies_btn = QPushButton("🍪 Получить Cookies")
        cookies_btn.setStyleSheet("background-color: #f39c12; color: white; padding: 12px;")
        cookies_btn.clicked.connect(self.parse_cookies)
        layout.addWidget(cookies_btn)

    def log(self, message):
        self.log_text.append(message)
        self.log_text.ensureCursorVisible()
        QApplication.processEvents()

    def select_excel(self):
        global excel_file
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите Excel", "", "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            excel_file = file_path
            self.file_label.setText(f"Выбран: {os.path.basename(file_path)}")
            self.log(f"Выбран файл: {file_path}")
        else:
            self.file_label.setText("Файл не выбран")

    def run_parsing(self):
        global excel_file
        if not excel_file:
            QMessageBox.warning(self, "Внимание", "Выберите Excel-файл!")
            return
        
        self.parse_btn.setEnabled(False)
        self.parse_btn.setText("Обработка...")
        self.log_text.clear()
        self.log("Запуск парсинга...")

        self.thread = ParsingThread(excel_file)
        self.thread.finished.connect(self.on_finished)
        self.thread.error.connect(self.on_error)
        self.thread.start()

    def on_finished(self, message):
        self.parse_btn.setEnabled(True)
        self.parse_btn.setText("🚀 Запустить парсинг")
        self.progress_bar.setVisible(False)
        self.log(message)
        QMessageBox.information(self, "Готово", message)

    def on_error(self, error_msg):
        self.parse_btn.setEnabled(True)
        self.parse_btn.setText("🚀 Запустить парсинг")
        self.log(f"Ошибка: {error_msg}")
        QMessageBox.critical(self, "Ошибка", error_msg)


    def open_dot_parsing(self):
        dialog = DotParsingDialog(self)
        dialog.exec_()

    def parse_cookies(self):
        try:
            # Твоя функция
            from cookies import parse_cookies
            success = parse_cookies()
            if success:
                self.log("Cookies сохранены")
                QMessageBox.information(self, "Успех", "Cookies сохранены!")
            else:
                self.log("Cookies не сохранены")
        except Exception as e:
            self.log(f"Ошибка cookies: {e}")
            QMessageBox.critical(self, "Ошибка", str(e))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())