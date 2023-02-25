import glob
import time
from pathlib import Path

import pandas as pd
from PyPDF3 import PdfFileReader, PdfFileWriter
from PyPDF3.pdf import PageObject


def create_list_barcode():
    """Формирует список штрихкодов для печати"""
    # Читаем в файле excel нужные столбцы
    excel_data = pd.read_excel('delivery.xlsx')
    data = pd.DataFrame(excel_data, columns=['article', 'quantity'])

    lenght_list = len(data['article'].to_list())
    # Список только нужных наименований
    key_list = data['article'].to_list()
    # Количество только нужных наименований
    value_list = data['quantity'].to_list()
    # Свисок всех наименований
    pdf_filenames = glob.glob("tickets/*")
    # Формируем список для печати
    new_list = []
    for i in range(lenght_list):
        for j in pdf_filenames:
            if str(key_list[i]) == str(Path(j).stem):
                while value_list[i] > 0:
                    new_list.append(j)
                    value_list[i] -= 1
    return(new_list)


def print_barcode_to_pdf():
    """Создает pdf файл для печати"""
    pdf_filenames = create_list_barcode()
    input1 = PdfFileReader(open(pdf_filenames[0], "rb"), strict=False)
    page1 = input1.getPage(0)
    total_width = max([page1.mediaBox.upperRight[0]*3])
    total_height = max([page1.mediaBox.upperRight[1]*6])
    output = PdfFileWriter()


    file_name = "barcods" + time.strftime("%Y%m%d") + ".pdf"
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


if __name__ == "__main__":
    print_barcode_to_pdf()
