# -*- coding: UTF-8 -*-  

# 扯淡的代码
# 交流：lzhwang@whu.edu.cn

import urllib2
import cookielib
import hashlib
import urllib
import re
import xlwt
def get_md5_value(src):
	# 登陆时md5计算
    myMd5 = hashlib.md5()
    myMd5.update(src)
    myMd5_Digest = myMd5.hexdigest()
    return myMd5_Digest
def getCsrfToken(content):
	# 获取查询成绩必须的csrfToken
	state = '''csrftoken=.{36}''' #正则
	titles = re.findall(state,content)
	try:
		print titles[0][10:]
	except IndexError:
		# 简单的一个判断,没有再花时间去区分验证码错or学号密码错
		# 因为要区分的话还要写正则 烦
		print '登陆信息有误，检查学号、密码、验证码'
		exit()
	return titles[0][10:]
def getScoreInfo(content):
	# 到成绩页面拉取成绩信息
	state = r'<tr null>(.*?)</tr>' #分课程在<tr null>标签内
	r = re.findall(state, content, re.S|re.M)
	result = []
	for lesson in r:
		stateL = r'<td>(.*?)</td>' # 课程信息在<td>标签里
		les = re.findall(stateL, lesson, re.S|re.M)
		lesson = []
		for i in range(0,10):
			lesson.append(les[i].encode("UTF-8"))
		result.append(lesson)
	return result
	# 返回值:result是一个二维数组
	# 其中的每个元素lesson都是每门课的具体信息构成的一维数组
	# 需要说明的是，lesson的每一个Index对应的信息分别为
	# 0 课头
	# 1 课程名
	# 2 类型
	# 3 学分
	# 4 教师
	# 5 开课学院
	# 6 学习类型
	# 7 学年
	# 8 学期
	# 9 成绩

def writeExcel(result):
	# 写Excel的方法，需要用到第三方库xlwt
	f = xlwt.Workbook(encoding = 'utf-8')
	sheet1 = f.add_sheet(u'成绩表',cell_overwrite_ok=True)
	row = 0
	for lesson in result:
		col = 0
		for info in lesson:
			sheet1.write(row,col, info, xlwt.Style.easyxf())
			col = col + 1
		row = row + 1
	f.save('score.xls')

class Splider(object):
	# 初始化，需要传入学号和密码
	# 初始化时还会要求输入验证码
	# 初始化后即与教务平台保持会话
	def __init__(self, studentnum, passwd):
		super(Splider, self).__init__()
		self.studentnum =studentnum
		self.passwd = passwd

		cookie = cookielib.CookieJar()
		# 利用urllib2库的HTTPCookieProcessor对象来创建cookie处理器
		handler=urllib2.HTTPCookieProcessor(cookie)
		# 通过handler来构建opener
		self.opener = urllib2.build_opener(handler)
		# opener存储的cookie始终保持一致

		imageUrl = 'http://210.42.121.133/servlet/GenImg'
		# 获取验证码，并要求手动输入
		req = urllib2.Request(imageUrl)
		content = self.opener.open(req).read()
		f = open('checkcode.jpg', 'wb')
		f.write(content)
		f.close()

		self.checkCode = raw_input(u"查看当前文件夹的checkcode.jpg文件，输入验证码")
	def run(self):
		req = urllib2.Request('http://210.42.121.133/servlet/Login')
		data = {'id':self.studentnum,'pwd':get_md5_value(self.passwd),'xdvfb':self.checkCode}
		data = urllib.urlencode(data)
		response = self.opener.open(req, data) # Login
		csrf = getCsrfToken(response.read())
		scoreUrl = 'http://210.42.121.133/servlet/Svlt_QueryStuScore?csrftoken='+csrf+'&year=0&term='
		req = urllib2.Request(scoreUrl)
		response = self.opener.open(req) # 拉取成绩页面
		return getScoreInfo(response.read()) # 返回前述二维数组

		
if __name__ == "__main__":
	studentnum = raw_input(u'学号')
	passwd = raw_input(u'密码')
	s = Splider(studentnum, passwd)
	result = s.run()
	writeExcel(result)
	print u"输出结果已存放在当前文件夹的score.xls文件中"