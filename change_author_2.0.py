# -*-config:utf-8-*-
"""
--------------------------------
File Name:        change_author_2.0.py
Author:           ZZQ
Date:             2023-06-12  8:55
--------------------------------
Change Activity:
                  2023-06-12  8:55
--------------------------------
"""
__author__ = 'ZZQ'
__version__ = "2.0"

import os
from docx import Document
from openpyxl import load_workbook


def word_author(dir_path, name, i):
	"""
	修改后缀名为.docx的Word文档的作者信息
	:param dir_path:
	:param name:
	:param i:
	:return:
	"""
	for filename in os.listdir(dir_path):
		if filename.endswith('.docx') and not filename.startswith('~$'):
			filepath = os.path.join(dir_path, filename)
			document = Document(filepath)
			core_properties = document.core_properties
			core_properties.author = name
			core_properties.last_modified_by = name
			document.save(filepath)
			i += 1
	return dir_path, i


def excel_author(dir_path, name, i):
	"""
	修改后缀名为.xlsx的Excel电子表格文档的作者信息
	:param dir_path:
	:param name:
	:param i:
	:return:
	"""
	for filename in os.listdir(dir_path):
		if filename.endswith('.xlsx') and not filename.startswith('~$'):
			filepath = os.path.join(dir_path, filename)

			# Load the workbook and remove the author property
			wb = load_workbook(filepath)
			wb.properties.creator = name
			wb.properties.last_modified_by = name

			# Save the changes to the workbook
			wb.save(filepath)
			i += 1

	return dir_path, i


def if_path(dir_path):
	"""
	判断用户输入的路径是否合法
	:param dir_path:
	:return: dir_path
	:return: name
	"""
	if os.path.isdir(dir_path) or dir_path == '':
        if dir_path == '':
            dir_path = './'
		dir_path = dir_path.replace('\\', '/')
		print('请输入你想要的作者姓名：')
		name = input()
		execute(dir_path, name)
	# elif 
		# print('请输入你想要的作者姓名：')
		# name = input()
		# execute(dir_path, name)
	else:
		print(f'您输入的路径有误，程序将退出！')


def execute(dir_path, name):
	"""
	调用处理代码入口
	:param dir_path:
	:param name:
	:return:
	"""
	while 1:
		selector = input('请选择欲处理的文档类型：1、Word  2、Excel  0、退出：')
		if selector == '1':
			count = 0
			word_path, count = word_author(dir_path, name, count)
			if count > 0:
				print(f'已处理指定目录  {word_path}  中 {count} 个".docx"文档的作者信息！')
			else:
				print(f'指定目录  {dir_path}  中没有要处理的文档！')
		elif selector == '2':
			count = 0
			excel_path, count = excel_author(dir_path, name, count)
			if count > 0:
				print(f'已处理指定目录  {excel_path}  中 {count} 个".xlsx"文档的作者信息！')
			else:
				print(f'指定目录  {dir_path}  中没有要处理的".xlsx"文档！')
		elif selector == '0':
			input('按任意键结束......')
			break
		else:
			print('选择无效，请重新选择！')


def main():
	"""
	程序入口
	:return:
	"""
	print('请输入文档所在的绝对路径后回车。')
	print('例如：C:/Users/ZZQ/Desktop/_docx')
	print('若直接回车，则默认指定当前目录。')
	dir_path = input()
	if_path(dir_path)


if __name__ == '__main__':
	main()