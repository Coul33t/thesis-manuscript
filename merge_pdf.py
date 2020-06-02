'''
This script is used to merge the PhD thesis and the cover (both PDFs).
'''

from PyPDF2 import PdfFileMerger

manuscript = open('main.pdf', 'rb')
cover_and_back = open('MathStic-LMU.pdf', 'rb')

merger = PdfFileMerger()

merger.append(fileobj=cover_and_back, pages=(0, 2))
merger.append(fileobj=manuscript)
merger.append(fileobj=cover_and_back, pages=(2, 3))

output = open("thesis_merged_with_cover.pdf", "wb")
merger.write(output)

manuscript.close()
cover_and_back.close()
output.close()