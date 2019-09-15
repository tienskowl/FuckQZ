#!/usr/bin/python3
# By 康康
import json
import datetime
import os

import openpyxl
from prettytable import PrettyTable, RANDOM
import requests


class JWXT:
	def __init__(self, acount, pwd):
		self.url = 'http://jwxt.upc.edu.cn/app.do?'
		self.header = {
			'User-Agent': 'Mozilla/5.0 (Linux; U; Mobile; Android 6.0.1;C107-9 Build/FRF91 )',
			'Referer': 'http://www.baidu.com',
			'accept-encoding': 'gzip, deflate, br',
			'accept-language': 'zh-CN,zh-TW;q=0.8,zh;q=0.6,en;q=0.4,ja;q=0.2',
			'cache-control': 'max-age=0'
		}
		self.number = acount
		self.pwd = pwd
		self.ss = self.login()
	
	def login(self):
		params = {
			"method": "authUser",
			"xh": self.number,
			"pwd": self.pwd
		}
		session = requests.session()
		req = session.get(self.url, params = params, timeout = 5, headers = self.header)
		s = json.loads(req.text)
		print(s['msg'])
		self.header['token'] = s['token']
		return session
	
	def getKbcxAzc(self, zc):
		params = {
			"method": "getKbcxAzc",
			# "xnxqid": s['xnxqh'], 选择学期，默认为当前学期
			"zc": zc,
			"xh": self.number
		}
		req = self.ss.get(self.url, params = params, headers = self.header)
		# print(req.text)
		return req.text
	
	def timeTrans(self, time):
		index = int((int(time[2] + int(time[1]) * 10) - 1) / 2)
		icstime = [['080000', '095000'], ['101000', '120000'], ['140000', '155000'], ['161000', '180000'],
		           ['190000', '205000'], ['210000', '225000']]
		return icstime[index]
	
	def create_ics(self, f):
		global date
		global n
		for week in range(1, 20):
			courses = json.loads(self.getKbcxAzc(week))
			for index, course in enumerate(courses):
				if course is None:
					break
				day = (date + datetime.timedelta(days = int(course['kcsj'][0]))).strftime('%Y%m%d')
				hour = self.timeTrans(course['kcsj'])
				message = '''BEGIN:VEVENT
SUMMARY:%s
DTSTART;TZID="UTC+08:00";VALUE=DATE-TIME:%sT%s
DTEND;TZID="UTC+08:00";VALUE=DATE-TIME:%sT%s
LOCATION:%s--%s
END:VEVENT\n''' % (
					str(n) + course['kcmc'], day, hour[0], day, hour[1], course['jsmc'], course['jsxm'])
				f.write(message)
			n += 1
			date += datetime.timedelta(days = 7)
			print(date)
	
	def getCjcx(self):
		params = {
			"method": "getCjcx",
			"xh": self.number,
			"xnxqid": ""
		}
		req = requests.get(self.url, params = params, headers = self.header)
		scores = json.loads(req.text)
		table = PrettyTable(["序号", "日期", "名称", "成绩", "学分", "类型", "考试性质"])
		# 创建excel表
		excel = openpyxl.Workbook()
		sheet = excel.active
		sheet.append(["序号", "日期", "名称", "成绩", "学分", "类型", "考试性质"])
		for i, score in enumerate(scores):
			sheet.append([i, score['xqmc'], score['kcmc'], score['zcj'], score['xf'], score['kclbmc'], score['ksxzmc']])
			table.add_row(
				[i, score['xqmc'], score['kcmc'], score['zcj'], score['xf'], score['kclbmc'], score['ksxzmc']])
		print(table)
		filename = 'cj' + self.number + '.xlsx'
		excel.save(filename)
		os.system("pause")


number = input('请输入学号')
pwd = input('请输入教务系统密码')
jw = JWXT(number, pwd)
print('token有过期机制，过期请重新打开程序')
option = input('选择操作：\n 1.生成课表 \n 2.查询成绩')
if option == '1':
	dat = input('请输入学期开始时间，2019-2020-1开始时间为2019 9 8')  # 获取学期第一天
	n = 1  # 周数计数
	tmp = dat.split(' ')
	date = datetime.datetime(int(tmp[0]), int(tmp[1]), int(tmp[2]))
	f = open('kb' + number + '.ics', 'w', encoding = 'utf-8')
	f.write(u"BEGIN:VCALENDAR\nVERSION:2.0\n")
	jw.create_ics(f)
	f.write(u"END:VCALENDAR")
	f.close()
else:
	jw.getCjcx()
