from PyPDF3 import PdfFileWriter, PdfFileReader
from PyPDF3.pdf import PageObject
import glob

pdf_filenames = glob.glob("tickets/*")

input1 = PdfFileReader(open(pdf_filenames[0], "rb"), strict=False)
input2 = PdfFileReader(open(pdf_filenames[1], "rb"), strict=False)
input3 = PdfFileReader(open(pdf_filenames[2], "rb"), strict=False)

page1 = input1.getPage(0)
page2 = input2.getPage(0)
page3 = input3.getPage(0)

total_width = max([page1.mediaBox.upperRight[0]*3])
total_height = max([page1.mediaBox.upperRight[1]*6])

new_page = PageObject.createBlankPage("result.pdf", total_width, total_height)
new_page_2 = PageObject.createBlankPage("result.pdf", total_width, total_height)

# Add first page at the 0,0 position
new_page.mergeTranslatedPage(page1, 0, 0)
# Add second page with moving along the axis x
new_page_2.mergeTranslatedPage(page2, page1.mediaBox.upperRight[0], page1.mediaBox.upperRight[1]*0)
new_page_2.mergeTranslatedPage(page3, page1.mediaBox.upperRight[0]*2, 0)

output = PdfFileWriter()
output.addPage(new_page)
output.addPage(new_page_2)
output.write(open("result.pdf", "wb"))