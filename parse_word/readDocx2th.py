import docx
import os
import shutil
from win32com import client
import re
from bs4 import BeautifulSoup
import pandas as pd


style = {
	'0': 'single',
	'1': 'multiple',
	'2': 'judgment',
	'3': 'fillin',
	'4': 'subjective',
	'5': 'english'
}

def read4word(path, html_path, main_txt_path):
	# for i in os.listdir(main_txt_path):
	# 	if os.path.isdir(os.path.join(main_txt_path, i)):
	# 		for file in os.listdir(os.path.join(main_txt_path, i)):
	# 			os.remove(os.path.join(main_txt_path, i, file))

	doc = docx.Document(path)
	end_para_no, QueStyle_para_no = read_doc4para_no(doc)
	print(QueStyle_para_no, end_para_no)
	word2html(path, html_path, end_para_no, QueStyle_para_no, main_txt_path)


def read_doc4para_no(doc):
	i = 1
	end_para_no = []
	QueStyle_para_no = []

	para = doc.paragraphs
	for p in range(len(para)):
		if re.compile("单选题示例").match(para[p].text):
			QueStyle_para_no.append(i)
		if re.compile("多选题示例").match(para[p].text):
			QueStyle_para_no.append(i)
		if re.compile("判断题示例").match(para[p].text):
			QueStyle_para_no.append(i)
		if re.compile("填空题示例").match(para[p].text):
			QueStyle_para_no.append(i)
		if re.compile("主观题示例").match(para[p].text):
			QueStyle_para_no.append(i)
		if re.compile("英语题示例").match(para[p].text):
			QueStyle_para_no.append(i)
		if para[p].text == "【结束】":
			end_para_no.append(i)
		i = i+1
	return end_para_no, QueStyle_para_no


def word2html(path, html_path, end_para_no, QueStyle_para_no, main_txt_path):
	Questions = []
	single = {}
	multiple = {}
	judgment = {}
	fillin = {}
	subjective = {}
	english = {}
	word = client.Dispatch('Word.Application')
	word.Visible = 0
	word.DisplayAlerts = 0
	doc = word.Documents.Open(path)
	doc.SaveAs(html_path + '.html', 10)  # 选用 wdFormatFilteredHTML的话公式图片将存储为gif格式
	doc.Close()
	word.Quit()
	file_path = html_path+'.html'
	htmlfile = open(file_path, 'r', encoding='gb2312')
	htmlhandle = htmlfile.read()
	soup = BeautifulSoup(htmlhandle, 'lxml')
	r1 = r'<span .*?>【题文】\S+</span>'
	r2 = r'<span .*?>【题文】</span>'
	r3 = r'<span .*?>【答案】\S+</span>'
	r4 = r'<span .*?>【答案】</span>'
	r5 = r'<span .*?>【解析】\S+</span>'
	r6 = r'<span .*?>【解析】</span>'
	style_no = 0
	for p in range(len(QueStyle_para_no) - 1):
		i = 1
		question = ""
		answer = ""
		analysis = ""
		for item in soup.find_all('p'):
			if QueStyle_para_no[p] <= i < QueStyle_para_no[p + 1]:
				if re.match(r"【题文】", item.text) is not None:
					line = str(item)
					if re.search(r1, line):
						line = line.replace("【题文】", "")
					if re.search(r2, line):
						m = re.search(r2, line)
						line = line.replace(m.group(0), "")
					question += line
				if re.match(r"【答案】", item.text)is not None:
					line = str(item)
					if re.search(r3, line):
						line = line.replace("【答案】", "")
					if re.search(r4, line):
						m = re.search(r4, line)
						line = line.replace(m.group(0), "")
					answer += line
				if re.match(r"【解析】", item.text)is not None:
					line = str(item)
					if re.search(r5, line):
						line = line.replace("【解析】", "")
					if re.search(r6, line):
						m = re.search(r6, line)
						line = line.replace(m.group(0), "")
					analysis += line
				if re.match(r"【结束】", item.text):
					mm = {}
					mm['type'] = style['%s' % str(p)]
					mm['question'] = question
					mm['answer'] = answer
					mm['analysis'] = analysis
					print(mm)
					Questions.append(mm)
					question = ""
					answer = ""
					analysis = ""
					# tihao += 1
			i += 1
		style_no = p
	question = ""
	answer = ""
	analysis = ""
	i = 1
	for item in soup.find_all('p'):
		if QueStyle_para_no[-1] <= i <= end_para_no[-1]:
			if re.compile(r"【题文】").match(item.text)is not None:
				line = str(item)
				if re.search(r1, line):
					line = line.replace("【题文】", "")
				if re.search(r2, line):
					m = re.search(r2, line)
					line = line.replace(m.group(0), "")
				question += line
			if re.compile(r"【答案】").match(item.text)is not None:
				line = str(item)
				if re.search(r3, line):
					line = line.replace("【答案】", "")
				if re.search(r4, line):
					m = re.search(r4, line)
					line = line.replace(m.group(0), "")
				answer += line
			if re.compile(r"【解析】").match(item.text)is not None:
				line = str(item)
				if re.search(r5, line):
					line = line.replace("【解析】", "")
				if re.search(r6, line):
					m = re.search(r6, line)
					line = line.replace(m.group(0), "")
				analysis += line
			if item.text == "【结束】":
				mm = {}
				mm['type'] = style['%s' % str(style_no+1)]
				mm['question'] = question
				mm['answer'] = answer
				mm['analysis'] = analysis
				print(mm)
				Questions.append(mm)
		i += 1
	# print(Questions)


if __name__ == '__main__':
	path = 'C:\\Users\\j20687\\Desktop\\demo.docx'
	txt_path = 'C:\\Users\\j20687\\Desktop\\HTML.txt'
	main_txt_path = 'C:\\Users\\j20687\\Desktop\\Main'
	html_path = 'C:\\Users\\j20687\\Desktop\\HTML'
	read4word(path, html_path, main_txt_path)






