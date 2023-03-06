import glob
import pandas as pd
import sys
import time
from pathlib import Path
from PyQt5.QtWidgets import (QApplication, QFileDialog, QGridLayout, 
                             QHBoxLayout, QMainWindow, QLabel,
                             QLineEdit, QPushButton, QStackedLayout,
                             QVBoxLayout, QWidget)
from PyPDF3 import PdfFileReader, PdfFileWriter
from PyPDF3.pdf import PageObject


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Innotrade")
        pagelayout = QVBoxLayout()
        button_layout = QHBoxLayout()
        self.stacklayout = QStackedLayout()
        pagelayout.addLayout(button_layout)
        pagelayout.addLayout(self.stacklayout)
        btn = QPushButton("сформировать файл PDF")
        btn.pressed.connect(self.print_barcode_to_pdf)
        button_layout.addWidget(btn)
        widget = QWidget()
        widget.setLayout(pagelayout)
        self.setCentralWidget(widget)


    def choseFolderBarcodes(self):
        """Выбираем папку где лежать все штрихкоды"""
        layout = QGridLayout()
        self.setLayout(layout)
        # directory selection
        dir_btn = QPushButton('Browse')
        self.dir_name_edit = QLineEdit()
        layout.addWidget(QLabel('label'), 0, 0)
        layout.addWidget(self.dir_name_edit, 1, 1)
        #layout.addWidget(dir_btn, 1, 2)
        #self.show()
        dir_name = QFileDialog.getExistingDirectory(
            self,
            "Выберите папку со штрихкодами",
            "D:\\"
        )
        if dir_name:
            path = Path(dir_name)
            self.dir_name_edit.setText(str(path))
        return(str(dir_name) + '/*')


    def choseFolderToSave(self):
        """Выбираем папку куда сохранять итоговый файл"""
        layout = QGridLayout()
        self.setLayout(layout)
        # directory selection
        dir_btn = QPushButton('Browse')
        self.dir_name_edit = QLineEdit()
        layout.addWidget(QLabel('label'), 0, 0)
        layout.addWidget(self.dir_name_edit, 1, 1)
        #layout.addWidget(dir_btn, 1, 2)
        #self.show()
        dir_name = QFileDialog.getExistingDirectory(
            self,
            "Выберите папку для сохранения файла",
            "D:\\"
        )
        if dir_name:
            path = Path(dir_name)
            self.dir_name_edit.setText(str(path))
        return(str(dir_name) + '/')


    def choseExcelFile(self):
        """Выбираем Excel файл с заказанными позициями"""
        layout1 = QGridLayout()
        self.setLayout(layout1)
        #file_browse = QPushButton('Путь до Excel файла')
        self.filename_edit = QLineEdit()
        layout1.addWidget(QLabel('label_2'), 1, 1)
        layout1.addWidget(self.filename_edit, 0, 1)
        #layout.addWidget(file_browse, 0 ,2)
        self.show()
        filename, ok = QFileDialog.getOpenFileName(
            self,
            "Выберите Excel файл", 
            "D:\\", 
            "Files (*.xls *.xlsx)"
        )
        if filename:
            path = Path(filename)
            self.filename_edit.setText(str(path))
        return(str(filename))


    def create_list_barcode(self):
        """Формирует список штрихкодов для печати"""
        # Читаем в файле excel нужные столбцы
        excel_data = pd.read_excel(self.choseExcelFile())
        data = pd.DataFrame(excel_data, columns=['article', 'quantity'])

        lenght_list = len(data['article'].to_list())
        # Список только нужных наименований
        key_list = data['article'].to_list()
        # Количество только нужных наименований
        value_list = data['quantity'].to_list()
        # Свисок всех наименований
        pdf_filenames = glob.glob(self.choseFolderBarcodes())
        # Формируем список для печати
        new_list = []
        for i in range(lenght_list):
            for j in pdf_filenames:
                if str(key_list[i]) == str(Path(j).stem):
                    while value_list[i] > 0:
                        new_list.append(j)
                        value_list[i] -= 1
        return(new_list)


    def print_barcode_to_pdf(self):
        """Создает pdf файл для печати"""
        pdf_filenames = self.create_list_barcode()
        input1 = PdfFileReader(open(pdf_filenames[0], "rb"), strict=False)
        page1 = input1.getPage(0)
        total_width = max([page1.mediaBox.upperRight[0]*3])
        total_height = max([page1.mediaBox.upperRight[1]*6])
        output = PdfFileWriter()
        file_name = f'{str(self.choseFolderToSave())}barcodes {time.strftime("%Y-%m-%d")}.pdf'
        new_page = PageObject.createBlankPage(file_name, total_width, total_height)
        new_page.mergeTranslatedPage(page1, 0, 0)
        output.addPage(new_page)
        page_amount = (len(pdf_filenames) // 18) + 1
        pages_names = []
    
        for p in range(1, page_amount):
            p = PageObject.createBlankPage(file_name, total_width, total_height)
            output.addPage(p)
            pages_names.append(p)
        new_page.mergeTranslatedPage(page1, 0, 0)

        for i in range(1, len(pdf_filenames)):
            m = i // 18
            n = (i // 3) - 6*m
            k = i % 3
            if i < 18:
                new_page.mergeTranslatedPage(
                (PdfFileReader(open(pdf_filenames[i], "rb"),
                           strict=False)).getPage(0),
                (PdfFileReader(open(pdf_filenames[0], "rb"),
                           strict=False)).getPage(0).mediaBox.upperRight[0]*k, 
                (PdfFileReader(open(pdf_filenames[0], "rb"),
                           strict=False)).getPage(0).mediaBox.upperRight[1]*n)

            elif i >= 18:
                (pages_names[m-1]).mergeTranslatedPage(
                (PdfFileReader(open(pdf_filenames[i], "rb"),
                           strict=False)).getPage(0),
                (PdfFileReader(open(pdf_filenames[0], "rb"),
                           strict=False)).getPage(0).mediaBox.upperRight[0]*k, 
                (PdfFileReader(open(pdf_filenames[0], "rb"),
                           strict=False)).getPage(0).mediaBox.upperRight[1]*n)
        output.write(open(file_name, "wb"))
     

if __name__ == '__main__':
    import sys
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())
