# -*- coding: utf-8 -*-
__author__ = 'lyc'

from xlrd import open_workbook
from xlutils.copy import copy
from xlwt import *
import datetime
import time
import math
import json


class MyHandle(object):
	input_file = u"data/加班申请.xls"
	output_file = u"output.xls"

	sheet_name = u'原始'
	top_title_row_num = 1

	idx_confirm_st     = 0
	idx_confirm_re     = 1
	idx_pid            = 2
	idx_pname          = 3

	idx_overtime_type  = 5
	idx_overtime_start = 6
	idx_overtime_end   = 7
	idx_overtime_cal   = 8

	idx_pdepart        = 9
	idx_output         = 10
	idx_output_comment = 11
	idx_output_ex_time_comment = 12

	confirm_st_complete = u'完成'
	confirm_re_accept   = u'同意'

	department_name_design = u'设计部'
	department_name_customer = u'客服部'

	workday_overtime = 0
	weekend_overtime = 1
	OVER_TIME_TYPE_MAP = {u'工作日加班【正常工作时间段外】': workday_overtime,
	                      u'休息日加班【正常工作时间段内】': weekend_overtime,}

	#### 注意 '00:00' 默认都是加到了明天，请格外注意 ####
	default_department = {workday_overtime:
		                      {'time_start_str': '09:30',
		                       'time_end_str': '18:00',
		                       'is_in': False},
	                      weekend_overtime:
		                      {'time_start_str': '09:30',
		                       'time_end_str': '17:30',
		                       'is_in': True}}

	design_department_t = ['00:00', '21:00', '22:00']
	design_department_hour = [8, 2, 4]

	customer_department = {workday_overtime:
		                      [{'time_start_str': '09:30',
		                       'time_end_str': '18:30',
		                       'is_in': False},
		                       {'time_start_str': '15:00',
		                       'time_end_str': '00:00',
		                       'is_in': False}],
	                       weekend_overtime:
		                      [{'time_start_str': '09:30',
		                       'time_end_str': '18:30',
		                       'is_in': True},
		                       {'time_start_str': '15:00',
		                       'time_end_str': '00:00',
		                       'is_in': True}]}

	string_id = {'overtime_start_wrong_when_in': u'加班开始时间有误，依正确时间计算(之内)',
	             'overtime_end_wrong_when_in': u'加班结束时间有误，依正确时间计算(之内)',
	             'overtime_start_wrong_when_out': u'加班开始时间有误，依正确时间计算(之外)',
	             'overtime_end_wrong_when_out': u'加班结束时间有误，依正确时间计算(之外)',
	             'overtime_wrong_when_design': u'加班时间有误，设计部',
	             'overtime_wrong_when_customer': u'加班时间有误，客服部',
	             'overtime_delta_over_oneday': u'加班时间超过24小时',
	             'overtime_delta_lt_zero': u'加班时间为负',
	             'not_approved': u'审批未完成或未同意',
	             'overtime_type_wrong': u'加班类型错误',
	             'program_error': u'程序错误'}

	color_red = 2
	color_magenta = 6

	def __init__(self):
		self._run()

	def _parse_config(self, config_file):
		with open(config_file) as f:
			jsonconfig = json.load(f)
		return jsonconfig

	def _run(self):
		r_wb = open_workbook(self.input_file, formatting_info=True, on_demand=True)
		r_sheet = r_wb.sheet_by_name(self.sheet_name)
		w_wb = copy(r_wb)
		w_sheet = w_wb.get_sheet(self.sheet_name)
		self._do_sheet(r_sheet, w_sheet)
		w_wb.save(self.output_file)
		print '-- Run success! --'

	def _write(self, w_sheet, row, col, data, color=None):
		if isinstance(data, list):
			real_data = ''
			for each in data:
				real_each = self.string_id.get(each) if col == self.idx_output_comment else each
				real_data += real_each
				real_data += u'\n'
		else:
			real_data = data

		if color is not None:
			style = XFStyle()
			pattern = Pattern()
			pattern.pattern = Pattern.SOLID_PATTERN
			pattern.pattern_fore_colour = color
			style.pattern = pattern
			w_sheet.write(row, col, real_data, style)
		else:
			w_sheet.write(row, col, real_data)

	def _do_sheet(self, r_sheet, w_sheet):
		for idx, row in enumerate(r_sheet.get_rows()):
			if idx < self.top_title_row_num:
				self._write(w_sheet, idx, self.idx_output, u'程序计算加班时间')
				self._write(w_sheet, idx, self.idx_output_comment, u'可疑情况的注释')
				self._write(w_sheet, idx, self.idx_output_ex_time_comment, u'规范化后的加班起始')
			else:
				self._do_row(idx, row, w_sheet)

	def _do_row(self, row_idx, row, w_sheet):
		if not self._check_available_row(row):
			self._out_put(w_sheet, row_idx, row, 0, ['not_approved'])
			return

		overtime_id = self.OVER_TIME_TYPE_MAP.get(row[self.idx_overtime_type].value, None)
		if overtime_id is None:
			self._out_put(w_sheet, row_idx, row, 0, ['overtime_type_wrong'])
			return

		overtime_start = self._transform_time_str(row[self.idx_overtime_start].value)
		overtime_end = self._transform_time_str(row[self.idx_overtime_end].value)

		if overtime_start >= overtime_end:
			self._out_put(w_sheet, row_idx, row, 0, ['overtime_delta_lt_zero'])
			return

		department_str = row[self.idx_pdepart].value
		ex_comment = []
		self._judge_mt_24(overtime_start, overtime_end, ex_comment)
		if department_str == self.department_name_design:
			# 设计部
			over_hours, comment, time_comment = self.do_design_cal(overtime_id, overtime_start, overtime_end)
		elif department_str == self.department_name_customer:
			# 客服部
			over_hours, comment, time_comment = self.do_customer_cal(overtime_id, overtime_start, overtime_end)
		else:
			# 其他部门
			over_hours, comment, time_comment = self.do_default_cal(overtime_id, overtime_start, overtime_end)
		comment.extend(ex_comment)
		self._out_put(w_sheet, row_idx, row, over_hours, comment, time_comment)

	def do_default_cal(self, overtime_id, overtime_start, overtime_end):
		conf = self.default_department.get(overtime_id, {})
		time_start_str = conf.get('time_start_str')
		time_end_str = conf.get('time_end_str')
		is_in = conf.get('is_in')
		assert time_start_str is not None or time_end_str is not None or is_in is not None
		return self._cal_result_comment(time_start_str, time_end_str, overtime_start, overtime_end, is_in)

	def do_design_cal(self, overtime_id, overtime_start, overtime_end):
		if overtime_id == self.workday_overtime:
			return self._get_desgin_time(overtime_start, overtime_end)
		elif overtime_id == self.weekend_overtime:
			return self.do_default_cal(overtime_id, overtime_start, overtime_end)
		else:
			return 0, ['overtime_type_wrong'], []

	def do_customer_cal(self, overtime_id, overtime_start, overtime_end):
		# return hours, comment
		comment = []
		info = self.customer_department.get(overtime_id)
		if not info:
			comment.append('program_error')
			return 0, comment, []
		result = []
		for each_couple in info:
			result.append(self._cal_result_comment(each_couple['time_start_str'], each_couple['time_end_str'],
			                                       overtime_start, overtime_end, is_in=each_couple['is_in']))
		return self._return_max_overtime(result)

	def _return_max_overtime(self, result):
		max_idx = 0
		max_hours = result[0][0]
		for idx, (hours, _, _) in enumerate(result):
			if hours > max_hours:
				max_idx = idx
				max_hours = hours
		return result[max_idx]

	def _get_desgin_time(self, overtime_start, overtime_end):
		tmp_flag_t = [self._gen_today_datetime(overtime_end, time_str, is_end=False) for time_str in self.design_department_t]
		if tmp_flag_t[0] == overtime_end:
			return self.design_department_hour[2], [], []
		if tmp_flag_t[0] < overtime_end < tmp_flag_t[1]:
			date_delta = overtime_end.date() - overtime_start.date()
			if date_delta.days == 0:
				return 0, ['overtime_wrong_when_design'], []
			else:
				return self.design_department_hour[0], [], []
		elif tmp_flag_t[1] <= overtime_end <= tmp_flag_t[2]:
			return self.design_department_hour[1], [], []
		else:
			return self.design_department_hour[2], [], []

	def _out_put(self, w_sheet, row_idx, row, over_hours, comment, ex_time_comment=[]):
		color = self._get_back_color(row, over_hours)
		self._write(w_sheet, row_idx, self.idx_output, over_hours, color)
		self._write(w_sheet, row_idx, self.idx_output_comment, comment)
		self._write(w_sheet, row_idx, self.idx_output_ex_time_comment, ex_time_comment)

	def _get_back_color(self, row, over_hours):
		if not over_hours:
			return self.color_red
		try:
			auto_result = row[self.idx_overtime_cal].value
			auto_result = auto_result[:-2]
			ori_hours = float(auto_result)
			if abs(ori_hours - round(over_hours, 2)) <= 0.0001:
				return None
			else:
				return self.color_magenta
		except:
			return None

	def _transform_time_str(self, time_str):
		# 2017-10-01 09:30
		return datetime.datetime.strptime(time_str, '%Y-%m-%d %H:%M')

	def _judge_mt_24(self, overtime_start, overtime_end, comment):
		original_time_delta = overtime_end - overtime_start
		if original_time_delta.days > 0:
			comment.append('overtime_delta_over_oneday')

	def _gen_today_datetime(self, overtime_datetime, time_str, is_end=True):
		overtime_start_date = overtime_datetime.date()
		tmp_time = time.strptime(time_str, '%H:%M')
		day_plus = False
		if time_str == '00:00' and is_end:
			day_plus = True
		ret = datetime.datetime(year=overtime_start_date.year, month=overtime_start_date.month, day=overtime_start_date.day,
		                               hour=tmp_time.tm_hour, minute=tmp_time.tm_min)
		if day_plus:
			ret += datetime.timedelta(days = 1)
		return ret

	def _cal_result_comment(self, time_start_str, time_end_str, overtime_start, overtime_end, is_in):
		'''
		:param time_start_str:
		:param time_end_str:
		:param overtime_start:
		:param overtime_end:
		:param is_in:
		:return:
		'''
		over_hours = 0
		comment = []
		time_comment = []
		time_start = self._gen_today_datetime(overtime_start, time_start_str)
		time_end = self._gen_today_datetime(overtime_start, time_end_str)

		if is_in:
			if overtime_start < time_start:
				comment.append('overtime_start_wrong_when_in')
				overtime_start = time_start
			if overtime_end > time_end:
				comment.append('overtime_end_wrong_when_in')
				overtime_end = time_end
			if comment:
				time_comment.append(self._gen_time_comment(overtime_start, overtime_end))
			over_hours = self._cal_time_delta(overtime_start, overtime_end)
		else:
			if overtime_start < time_start:
				if overtime_end <= time_start:
					over_hours = self._cal_time_delta(overtime_start, overtime_end)
				elif overtime_end <= time_end:
					comment.append('overtime_end_wrong_when_out')
					time_comment.append(self._gen_time_comment(overtime_start, time_start))
					over_hours = self._cal_time_delta(overtime_start, time_start)
				else:
					comment.append('overtime_start_wrong_when_out')
					comment.append('overtime_end_wrong_when_out')
					time_comment.append(self._gen_time_comment(overtime_start, time_start))
					time_comment.append(self._gen_time_comment(time_end, overtime_end))
					over_hours = self._cal_time_delta(overtime_start, time_start) + self._cal_time_delta(time_end, overtime_end)
			elif overtime_start < time_end:
				comment.append('overtime_start_wrong_when_out')
				if overtime_end <= time_end:
					comment.append('overtime_end_wrong_when_out')
					overtime_end = time_end
				time_comment.append(self._gen_time_comment(time_end, overtime_end))
				over_hours = self._cal_time_delta(time_end, overtime_end)
			else:
				over_hours = self._cal_time_delta(overtime_start, overtime_end)

		return over_hours, comment, time_comment

	def _gen_time_comment(self, overtime_start, overtime_end):
		return str(overtime_start) + ' - ' + str(overtime_end)

	def _cal_time_delta(self, start_t, end_t):
		if end_t <= start_t:
			return 0
		time_delta = end_t - start_t
		total_seconds = time_delta.total_seconds()
		hours = total_seconds / 3600.0
		return hours

	def _check_available_row(self, row):
		confirm_st = row[self.idx_confirm_st]
		confirm_re = row[self.idx_confirm_re]
		if confirm_st.value != self.confirm_st_complete or confirm_re.value != self.confirm_re_accept:
			return False
		return True


if __name__ == '__main__':
	hle = MyHandle()