# -*- coding: utf-8 -*-
__author__ = 'lyc'

from xlrd import open_workbook
from xlutils.copy import copy
import datetime
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
	idx_pdepart        = 4
	idx_overtime_type  = 5
	idx_overtime_start = 6
	idx_overtime_end   = 7
	idx_overtime_cal   = 8

	idx_output         = 9
	idx_output_comment = 10

	confirm_st_complete = u'完成'
	confirm_re_accept   = u'同意'

	depart_type_0 = 0
	depart_type_1 = 1
	DEPARTMENT_TYPE_MAP = {u'辅料采购': 0,
	                       u'林珊珊组': 1,}

	OVER_TIME_TYPE_MAP = {u'工作日加班【正常工作时间段外】': 0,
	                      u'休息日加班【正常工作时间段内】': 1,}

	def __init__(self, config_file):
		self.config = self._parse_config(config_file)
		if not self.config:
			print 'Error: config parse failed.'
			return
		self.input_file = self.config['input_file']
		self.out_path = self.config['output_file']
		self.sheet_name = self.config['sheet_name']
		self.top_title_row_num = self.config['top_title_row_num']

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
		w_wb.save(self.out_path)

	def _write(self, w_sheet, row, col, data):
		w_sheet.write(row, col, data)

	def _do_sheet(self, r_sheet, w_sheet):
		for idx, row in enumerate(r_sheet.get_rows()):
			if idx < self.top_title_row_num:
				continue
			self._do_row(idx, row, w_sheet)

	def _do_row(self, row_idx, row, w_sheet):
		if not self._check_available_row(row):
			self._write(w_sheet, row_idx, self.idx_output, u'not_available')
			return

		department_id = self.DEPARTMENT_TYPE_MAP.get(row[self.idx_pdepart], None)
		overtime_id = self.OVER_TIME_TYPE_MAP.get(row[self.idx_overtime_type], None)
		overtime_start = self._transform_time_str(row[self.idx_overtime_start].value)
		overtime_end = self._transform_time_str(row[self.idx_overtime_end].value)
		result = self._do_cal(department_id, overtime_id, overtime_start, overtime_end)
		self._write(w_sheet, row_idx, self.idx_output, result)

	def _do_cal(self, department_id, overtime_id, overtime_start, overtime_end):
		result = 0
		start_t = datetime.datetime(hour=8)
		end_t = datetime.datetime(hour=18)

		return self._cal_time_delta(overtime_start, overtime_end)

	def _transform_time_str(self, time_str):
		# 2017-10-01 09:30
		return datetime.datetime.strptime(time_str, '%Y-%m-%d %H:%M')

	def _cal_time_delta(self, time_start, time_end):
		time_delta = time_end - time_start
		total_seconds = time_delta.total_seconds()
		hours = math.ceil(total_seconds / 3600.0)
		return hours

	def _check_available_row(self, row):
		confirm_st = row[self.idx_confirm_st]
		confirm_re = row[self.idx_confirm_re]
		if confirm_st.value != self.confirm_st_complete or confirm_re.value != self.confirm_re_accept:
			return False
		return True


if __name__ == '__main__':
	config_file = 'my.conf'
	hle = MyHandle(config_file)