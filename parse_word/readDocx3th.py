import docx
from win32com import client
import re
from bs4 import BeautifulSoup
import os



'''
试卷按顺序的题型放入style中
'''
style = {
	'0': 'single',
	'1': 'multiple',
	'2': 'judgment',
	'3': 'fillin',
	'4': 'subjective',
	'5': 'english'
}

def read4word(file):
	'''
	:param file: 为传入的.docx文件对象,以二进制格式打开
	:return:
	'''
	doc = docx.Document(file)
	end_para_no, QueStyle_para_no = read_doc4para_no(doc)
	# print(QueStyle_para_no, end_para_no)
	word2html(file, end_para_no, QueStyle_para_no)


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


def word2html(file, end_para_no, QueStyle_para_no):
	Questions = []
	word = client.Dispatch('Word.Application')
	# 后台运行，不显示，不警告
	word.Visible = 0
	word.DisplayAlerts = 0

	doc = word.Documents.Open(file)
	doc.SaveAs('\\tmp\\html.html', 10)  # 选用 wdFormatFilteredHTML的话公式图片将存储为gif格式
	doc.Close()
	word.Quit()
	file_path = '\\tmp\\html.html'
	htmlfile = open(file_path, 'r', encoding='gb2312' or 'utf-8')
	htmlhandle = htmlfile.read()
	soup = BeautifulSoup(htmlhandle, 'lxml')
	style_no = 0
	patt = re.compile(r'<img.*?src=".*?\.(?:jpg|jpeg|gif|bmp|png)">|(<span).*?(>)(.*?)(</span>)', re.S)
	re_a = re.compile(r'<span>A\.(.*?)B', re.S | re.M)
	re_b = re.compile(r'B\.(.*?)C', re.S | re.M)
	re_b_judgement = re.compile(r'B\.(.*?)</p>', re.S | re.M)
	re_c = re.compile(r'C\.(.*?)D', re.S | re.M)
	re_d = re.compile(r'D\.(.*?)</p>', re.S | re.M)

	for p in range(len(QueStyle_para_no) - 1):
		i = 1
		question = ""
		option = {}
		answer = ""
		analysis = ""
		for item in soup.find_all('p'):
			para_str = ""
			if QueStyle_para_no[p] <= i < QueStyle_para_no[p + 1]:
				group = patt.findall(str(item))
				for g in group:
					span_str = g[0] + g[1] + g[2] + g[3]
					para_str += span_str
				para_str = "<p>" + para_str + "</p>"
				para_str = para_str.replace("\n", "")
				if "【题文】" in para_str:
					question = para_str.replace("<span>【题文】</span>", "")
				if "【选项】" in para_str:
					if style['%s' % str(p)] == 'judgment':
						opt = para_str.replace("<span>【选项】</span>", "")
						A = re_a.findall(opt)
						B = re_b_judgement.findall(opt)
						option['A'] = str(A[0]).replace('<span>', '').replace('</span>', '')
						option['B'] = str(B[0]).replace('<span>', '').replace('</span>', '')
					else:
						opt = para_str.replace("<span>【选项】</span>", "")
						A = re_a.findall(opt)
						B = re_b.findall(opt)
						C = re_c.findall(opt)
						D = re_d.findall(opt)
						option['A'] = str(A[0]).replace('<span>', '').replace('</span>', '')
						option['B'] = str(B[0]).replace('<span>', '').replace('</span>', '')
						option['C'] = str(C[0]).replace('<span>', '').replace('</span>', '')
						option['D'] = str(D[0]).replace('<span>', '').replace('</span>', '')
				if "【答案】" in para_str:
					answer = para_str.replace("<span>【答案】</span>", "")
				if "【解析】" in para_str:
					analysis = para_str.replace("<span>【解析】</span>", "")
				if "【结束】" in para_str:
					mm = {}
					if style['%s' % str(p)] == 'subjective':
						mm['type'] = style['%s' % str(p)]
						mm['question'] = question
						mm['answer'] = answer
						mm['analysis'] = analysis
						print(mm)
						Questions.append(mm)
						question = ""
						option = ""
						answer = ""
						analysis = ""
					else:
						mm['type'] = style['%s' % str(p)]
						mm['question'] = question
						mm['option'] = option
						mm['answer'] = answer
						mm['analysis'] = analysis
						print(mm)
						Questions.append(mm)
						question = ""
						option = ""
						answer = ""
						analysis = ""
			i += 1
		style_no = p
	question = ""
	option = {}
	answer = ""
	analysis = ""
	i = 1
	for item in soup.find_all('p'):
		para_str = ""
		if QueStyle_para_no[-1] <= i <= end_para_no[-1]:
			group = patt.findall(str(item))
			for g in group:
				span_str = g[0] + g[1] + g[2] + g[3]
				para_str += span_str
			para_str = "<p>" + para_str + "</p>"
			para_str = para_str.replace("\n", "")
			if "【题文】" in para_str:
				question = para_str.replace("<span>【题文】</span>", "")
			if "【选项】" in para_str:
				opt = para_str.replace("<span>【选项】</span>", "")
				A = re_a.findall(opt)
				B = re_b.findall(opt)
				C = re_c.findall(opt)
				D = re_d.findall(opt)
				option['A'] = str(A[0]).replace('<span>', '').replace('</span>', '')
				option['B'] = str(B[0]).replace('<span>', '').replace('</span>', '')
				option['C'] = str(C[0]).replace('<span>', '').replace('</span>', '')
				option['D'] = str(D[0]).replace('<span>', '').replace('</span>', '')
			if "【答案】" in para_str:
				answer = para_str.replace("<span>【答案】</span>", "")
			if "【解析】" in para_str:
				analysis = para_str.replace("<span>【解析】</span>", "")
			if "【结束】" in para_str:
				mm = {}
				if style['%s' % str(style_no+1)] == 'subjective' or 'english':
					mm['type'] = style['%s' % str(style_no+1)]
					mm['question'] = question
					mm['answer'] = answer
					mm['analysis'] = analysis
					print(mm)
					Questions.append(mm)
					question = ""
					option = ""
					answer = ""
					analysis = ""
				else:
					mm['type'] = style['%s' % str(style_no+1)]
					mm['question'] = question
					mm['option'] = option
					mm['answer'] = answer
					mm['analysis'] = analysis
					print(mm)
					Questions.append(mm)
					question = ""
					option = ""
					answer = ""
					analysis = ""
		i += 1
	# print(Questions)


if __name__ == '__main__':
	# path = 'C:\\Users\\j20687\\Desktop\\demo.docx'
	# html_path = 'C:\\Users\\j20687\\Desktop\\HTML'
	read4word(file)






