import pandas as pnd
from docx import Document
import time

excel_file = 'test.xlsx'
word_doc = Document('template.docx')

replaced_word = 'NONAME'
column_name = 'NAME'

data_frame = pnd.read_excel(excel_file)
excel_data = set(data_frame[column_name].tolist())

	
def find_word(value):
	for paragraph in word_doc.paragraphs:
		if replaced_word in paragraph.text:
			print(paragraph.text)
			paragraph.text = paragraph.text.replace(replaced_word, value)
			print(paragraph.text)
			file_name = value + '.docx'
			word_doc.save(file_name)


def result(key):
	find_word(key)


for key in excel_data:
    result(key)


for key in excel_data:
    print(key)
    result(key)
    print(key)


