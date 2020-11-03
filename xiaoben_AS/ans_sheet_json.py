

def pre_info():
	pre_info = {}
	height = 2101
	width = 1542
	mark_base64_str = ""
	target_base64_str = ""
	paper_num = 0
	sub_mark_num_vec = []
	sub_mark_num = 34
	question_type = 0
	return pre_info


def choice_info():
	choice_info = {}
	choice_answer_inf = {}
	choice_col_num = 3
	col_to_col_gap = 227
	horizon_gap = 16
	left_to_mark = [89, 511, 932]
	question_type = 1
	rect_height = 28
	rect_thick = 2
	rect_width = 37
	scale = 0.54474708171206221
	submark = [2, 5]
	top_to_mark = [30, 335]
	vertical_gap = 15
	return choice_info


def fillin_info():
	fillin_info = {}
	blankfillin_answer_inf = {
		"blankfillin_answer_cnt": 4,
		"blankfillin_answer_pattern": [4]
	}
	question_type = 2
	submark = [6, 9, 10, 13]
	return fillin_info


def eassy_info():
	eassy_info = {}
	question_type = 2
	submark = [6, 9, 10, 13]
	return eassy_info


def get_json():
	json_data = []
	# 总体数据
	pre_info = pre_info()
	# 选择题数据
	choice_info = choice_info()
	# 填空题数据
	fillin_info = fillin_info()
	# 简答题数据
	eassy_info = eassy_info()
	json_data.append(pre_info)
	json_data.append(choice_info)
	json_data.append(fillin_info)
	json_data.append(eassy_info)


