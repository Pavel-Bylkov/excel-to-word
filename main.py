from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (QApplication, QWidget, QLabel, QLineEdit, QPushButton,
                            QListWidget, QVBoxLayout, QHBoxLayout, QFileDialog, QMessageBox)
import os

from openpyxl import load_workbook
from docxtpl import DocxTemplate

from pprint import pprint

class MyWin(QWidget):
    def __init__(self):
        super().__init__()
        self.filename = ""
        self.template_name = ""
        self.path_result_dir = ""
        self.config()
        self.init_gui()
        self.connect()
        self.wb = None

    def config(self):
        self.setWindowTitle("Приложение")
        self.resize(500, 300)

    def init_gui(self):
        self.lb_input_file = QLabel("Файл с данными:")
        self.le_input_file = QLineEdit()
        self.btn_input_file = QPushButton("Выбор файла")

        self.lb_output_file = QLabel("Файл с шаблоном:")
        self.le_output_file = QLineEdit()
        self.btn_output_file = QPushButton("Выбор файла")

        self.lb_static = QLabel("Диапазон констант")
        self.le_static = QLineEdit()
        self.lb_dynamic = QLabel("Диапазон переменных")
        self.le_dynamic = QLineEdit()

        self.lb_result_dir = QLabel("Папка с результатами:")
        self.le_result_dir = QLineEdit()
        self.btn_result_dir = QPushButton("Выбор папки")

        self.btn_run = QPushButton("Запуск обработки")

        v_line = QVBoxLayout()
        row1 = QHBoxLayout()
        row1.addWidget(self.lb_input_file)
        row1.addWidget(self.le_input_file)
        row1.addWidget(self.btn_input_file)
        v_line.addLayout(row1)
        row1_2 = QHBoxLayout()
        row1_2.addWidget(self.lb_output_file)
        row1_2.addWidget(self.le_output_file)
        row1_2.addWidget(self.btn_output_file)
        v_line.addLayout(row1_2)
        v_line.addStretch(1)
        row2 = QHBoxLayout()
        row2.addWidget(self.lb_static)
        row2.addWidget(self.le_static)
        row2.addStretch(1)
        v_line.addLayout(row2)
        row3 = QHBoxLayout()
        row3.addWidget(self.lb_dynamic)
        row3.addWidget(self.le_dynamic)
        row3.addStretch(1)
        v_line.addLayout(row3)
        v_line.addStretch(1)
        row4 = QHBoxLayout()
        row4.addWidget(self.lb_result_dir)
        row4.addWidget(self.le_result_dir)
        row4.addWidget(self.btn_result_dir)
        v_line.addLayout(row4)
        v_line.addWidget(self.btn_run)
        self.setLayout(v_line)

    def connect(self):
        self.btn_input_file.clicked.connect(self.input_file)
        self.le_input_file.editingFinished.connect(self.set_filename)
        self.btn_run.clicked.connect(self.run)
        self.btn_output_file.clicked.connect(self.output_file)
        self.le_output_file.editingFinished.connect(self.set_template)
        self.btn_result_dir.clicked.connect(self.result_dir)
        self.le_result_dir.editingFinished.connect(self.set_result_dir)

    def input_file(self):
        self.filename = QFileDialog.getOpenFileName(self)[0]
        if self.filename:
            self.le_input_file.setText(self.filename)

    def output_file(self):
        self.template_name = QFileDialog.getOpenFileName(self)[0]
        if self.template_name:
            self.le_output_file.setText(self.template_name)

    def result_dir(self):
        self.path_result_dir = QFileDialog.getExistingDirectory(parent=self)
        if self.path_result_dir:
            self.le_result_dir.setText(self.path_result_dir)

    def set_filename(self):
        self.filename = self.le_input_file.text()
        if not os.path.exists(self.filename) or not os.path.isfile(self.filename):
            self.filename = ""
            self.le_input_file.setText("")
    
    def set_template(self):
        self.filename = self.le_input_file.text()
        if not os.path.exists(self.template_name) or not os.path.isfile(self.template_name):
            self.template_name = ""
            self.le_output_file.setText("")

    def set_result_dir(self):
        self.path_result_dir = self.le_result_dir.text()
        if not os.path.exists(self.path_result_dir):
            self.path_result_dir = ""
            self.le_result_dir.setText("")

    def run(self):
        if (self.filename and self.le_static.text() and
                self.le_output_file.text() and self.le_dynamic.text()
                and self.le_result_dir.text()):
            try:
                self.wb = load_workbook(filename=self.filename)
                sheet = self.wb.active
                static_cells = sheet[self.le_static.text()]
                dynamic_cells = sheet[self.le_dynamic.text()]
                self.static_dict = {}
                for row in static_cells:
                    self.static_dict[row[0].value] = row[1].value
                self.dynamic_dict = {}
                for row in dynamic_cells:
                    self.dynamic_dict[row[0].value] = [cell.value for cell in row[1:]]
            except:
                QMessageBox.warning(self, "Ошибка", "Ошибка чтения данных")
            self.processing()

    def processing(self):
        try:
            self.temlate = DocxTemplate(self.template_name)
            key = list(self.dynamic_dict.keys())[0]
            for i in range(len(self.dynamic_dict[key])):
                self.create_file(i)
        except:
            QMessageBox.warning(self, "Ошибка", "Ошибка чтения шаблона")

    def create_file(self, i):
        try:
            context = self.static_dict.copy()
            for key, value in self.dynamic_dict.items():
                context[key] = value[i]
            self.temlate.render(context)
            output_file_name = f"{self.le_output_file.text().split('.')[0].split(os.sep)[-1]}_{i}.docx"
            self.temlate.save(self.path_result_dir + os.sep + output_file_name)
        except:
            QMessageBox.warning(self, "Ошибка", "Ошибка создания файла")


def main():
    app = QApplication([])
    win = MyWin()
    win.show()
    app.exec()

main()