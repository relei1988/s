# -*- coding: utf-8 -*- 
import  xdrlib ,sys
import xlrd
import docx
import re
data = xlrd.open_workbook('SS.xlsx')

def ks(a,b):
	if a>b:
		c=a-b
		d="扩大"
		return c,d
	elif b>a:
		c=b-a
		d="缩小"
		return c,d
	else:
		c=""
		d="持平"
		return c,d	

def zj(a,b):
	if a>b:
		c=a-b
		d="增加"
		return d,c
	elif b>a:
		c=b-a
		d="减少"
		return d,c
	else:
		c=""
		d="持平"
		return d,c	

def zd(a,b):
	if a>b:
		c=a-b
		d="涨"
		return d,c
	elif b>a:
		c=b-a
		d="跌"
		return d,c
	else:
		c=""
		d="持平"
		return d,c		

def neipan():
	
	table = data.sheet_by_name(u'卓创价格')
	#table = data.sheets()[20]
	'''
	name_list=[]//获取列名
	name_list.extend(table.row_values(1))//添加到数组中
	'''
	for i in range(table.nrows):
		v = table.cell(i,0)
		a = str(v)
		if "xl"  not in a and i >10:
			break	
	#得到i-1是最后一行有效数据
	lastline=i-1
	lastweek = i-6
	taicang_last = str(table.cell(lastweek,2))[7:11]
	huanan_last=str(table.cell(lastweek,3))[7:11]
	hebei_last=str(table.cell(lastweek,1))[7:11]
	lunan_last=str(table.cell(lastweek,4))[7:11]


	hebei=str(table.cell(lastline,1))[7:11]
	taicang=str(table.cell(lastline,2))[7:11]
	huanan=str(table.cell(lastline,3))[7:11]
	lunan=str(table.cell(lastline,4))[7:11]
	eeds=str(table.cell(lastline,5))[7:11]

	a=int(taicang)
	b=int(taicang_last)
	ihun=int(huanan)
	ihunl=int(huanan_last)
	ihb=int(hebei)
	ihbl=int(hebei_last)
	iln=int(lunan)
	ilnl=int(lunan_last)

	c = a - b
	d = ihun-ihunl
	e = ihb-ihbl
	f = iln-ilnl

	if c >0:
		flag = "涨"
	else:
		flag= "跌"
		c=abs(c)

	if d >0:
		flag0 = "涨"
	else:
		flag0 = "跌"
		d=abs(d)
	if e >0:
		flag1 = "涨"
	else:
		flag1 = "跌"
		e=abs(e)		
	
	if f >0:
		flag2 = "涨"
	else:
		flag2 = "跌"
		f=abs(f)


	c1 = """周报:
	1、现货价格及成交情况:
	截至本周五，太仓现货收至%s，%s%s。华南%s，%s%s。
	西北南线%s，北线%s，河北石家庄%s，%s%s。鲁南%s,%s%s。

	"""%(taicang,flag,c,huanan,flag0,d,eeds,eeds,hebei,flag1,e,lunan,flag2,f)
	print c1
	#本周价差
	nmjs=float(str(table.cell(lastline,12))[7:])
	nmsd=float(str(table.cell(lastline,14))[7:])
	sdjs=float(str(table.cell(lastline,15))[7:])
	hbjs=float(str(table.cell(lastline,13))[7:])
	nmhb=float(str(table.cell(lastline,11))[7:])
	#上周价差
	nmjsl=float(str(table.cell(lastline-5,12))[7:])
	nmsdl=float(str(table.cell(lastline-5,14))[7:])
	sdjsl=float(str(table.cell(lastline-5,15))[7:])
	hbjsl=float(str(table.cell(lastline-5,13))[7:])
	nmhbl=float(str(table.cell(lastline-5,11))[7:])

	a,b=ks(nmjs,nmjsl)
	c,d=ks(nmsd,nmsdl)
	e,f=ks(sdjs,sdjsl)
	j,k=ks(hbjs,hbjsl)
	l,m=ks(nmhb,nmhbl)
	c2='''
	本周国内各地区间价差情况：内蒙-江苏价差%s，%s%s。内蒙-山东价差%s，%s%s。山东-江苏价差%s，%s%s。内蒙-河北价差%s，%s%s。河北-江苏价差%s，%s%s。
	'''	%(nmjs,b,a,nmsd,d,c,sdjs,f,e,nmhb,m,l,hbjs,k,j)
	print c2


def waipan():
	'''
	data = xlrd.open_workbook('SS.xlsx')
	'''
	table = data.sheet_by_name(u'国际价格')	
	for i in range(table.nrows):
		v = table.cell(i,0)
		a = str(v)
		if "xl"  not in a and i >150:
			break
		#此处i为最后一行数据
		crf = str(table.cell(i,3))[7:]
		crf_last = str(table.cell(i-5,3))[7:]
	dcrf=float(crf)
	dcrfl=float(crf_last)

	a = dcrf-dcrfl

	if a > 0:
		flag = "涨"
	else:
		flag= "跌"
		a=abs(a)

	c1="""2、外盘
	本周甲醇中国CFR收至%s美金，%s%s美金，
	"""%(dcrf,flag,a)
	print c1
	table = data.sheet_by_name(u'卓创价格')	
	for i in range(table.nrows):
		v = table.cell(i,19)
		a = str(v)
		if "num" not in a and i>250:
			break
	#i-1是最后一个华东进口利润的行数
	hd=float(str(table.cell(i-1,19))[7:])
	hdl=float(str(table.cell(i-6,19))[7:])


	'''
	hd = str(table.cell(i -1,19))
	hdl=str(table.cell(i-6,19))

	fhd= re.findall(r"\d+\.?\d*",hd)
	fhdl = re.findall(r"\d+\.?\d*",hdl)

	a = float(fhd[0])
	b = float(fhdl[0])
	div = a - b
	'''
	div = hd-hdl
	print """
	华东进口利润%s元/吨，较上周四%s元/吨。
	"""%(hd,div)

#华东进口利润??元/吨，亏损扩大??元/吨。
'''
	
def doc_ctrl():
	doc = docx.Document(ur'report.docx')
	print doc.inline_shapes()
'''
def shui():
	table = data.sheet_by_name(u'每日报价')
	for i in range(table.nrows):
		v=table.cell(i,16)
		a =str(v)
		if "empty" in a and i >700:
			break
	#i-1是最后一个交易		
	ma05= str(table.cell(i,16))
	ma05l=str(table.cell(i-5,16))
	ima05=float(re.findall(r"\d*",ma05)[7])
	ima05l = float(re.findall(r"\d*",ma05l)[7])
	pct=(ima05-ima05l)/ima05l*100
	ma09=str(table.cell(i,17))
	ma09l=str(table.cell(i-5,17))
	ima09=float(re.findall(r"\d*",ma09)[7])
	ima09l=float(re.findall(r"\d*",ma09l)[7])
	pct09=(ima09-ima09l)/ima09l*100
	a=ima05-ima05l
	b=ima09-ima09l
	if a>0:
		flag1="涨"
	else:
		flag1="跌"
		a=abs(a)
	if b>0:
		flag2="涨"
	else:
		flag2="跌"
		b=abs(b)
	shui=str(table.cell(i,21))

	ishui=int(re.findall(r"\d*",shui)[7])
	if ishui >0:
		flag3="升水"
	else:
		flag3="贴水"
	shuil=str(table.cell(i-5,21))
	ishuil=int(re.findall(r"\d*",shuil)[7])
	c=ishui-ishuil
	if c>0:
		flag4="扩大"
	else:
		flag4="缩小"
		c=abs(c)	
	gap=int(ima05-ima09)
	gapl=int(ima05l-ima09l)
	d = gap-gapl
	if d>0:
		flag5="扩大"
	else:
		flag5="缩小"
		d=abs(d)

	print """二、盘面升贴水
	截至本周五，ma05收至%s，%s%s点（%s %%），ma09收至%s，%s%s点(%s %%)
	甲醇现货对MA05%s%s，%s%s。MA05-MA09价差%s，%s%s
	"""%(ima05,flag1,a,pct,ima09,flag2,b,pct09,flag3,ishui,flag4,c,gap,flag5,d)


def kucun():
	table = data.sheet_by_name(u'卓创库存')
	for i in range(table.nrows):
		v=table.cell(i,5)
		a =str(v)
		if "empty" in a and i >100:
			break

	kc=float(str(table.cell(i,5))[7:])
	kcl=float(str(table.cell(i-1,5))[7:])
	a,b=zj(kc,kcl)
	jskc=float(str(table.cell(i,8))[7:])
	jskcl=float(str(table.cell(i-1,8))[7:])
	c,d=zj(jskc,jskcl)


	print """三、库存
	本周甲醇港口库存%s万吨，%s%s万吨。本周江苏库存%s万吨，%s%s万吨。
	根据卓创统计，本周太仓日均提货量()（上周日均提货量）。
	"""%(kc,a,b,jskc,c,d)

def gongji():
	table = data.sheet_by_name(u'卓创开工率')
	for i in range(table.nrows):
		v=table.cell(i,1)
		a =str(v)
		if "empty" in a:
			break
	rate=float(str(table.cell(i-1,1))[7:])*100
	ratel=float(str(table.cell(i-2,1))[7:])*100
	a,b=zj(rate,ratel)
	print '''四、供给
	国内：本周国内甲醇工厂开工负荷%s%%，%s%s%%。
	'''%(rate,a,b)


def xuqiu():
	table = data.sheet_by_name(u'卓创开工率')
	for i in range(table.nrows):
		v=table.cell(i,1)
		a =str(v)
		if "empty" in a:
			break
	fuhe=float(str(table.cell(i-1,9))[7:])*100
	fuhel=float(str(table.cell(i-2,9))[7:])*100
	a,b=zj(fuhe,fuhel)
	print '''五、需求
	1、烯烃下游
		本周国内mto/mtp装置负荷%s%%，%s%s%%。
	'''%(fuhe,a,b)
	table = data.sheet_by_name(u'甲醇与PP')
	for i in range(table.nrows):
		v=table.cell(i,3)
		a =str(v)
		if "empty" in a:
			break
	lirun = float(str(table.cell(i-1,3))[7:])
	lirunl = float(str(table.cell(i-6,3))[7:])
	lirun05 = float(str(table.cell(i-1,29))[7:])
	lirun05l = float(str(table.cell(i-6,29))[7:])
	lirun09 = float(str(table.cell(i-1,33))[7:])
	lirun09l = float(str(table.cell(i-6,33))[7:])	

	a,b=ks(lirun,lirunl)
	c,d=ks(lirun05,lirun05l)
	e,f=ks(lirun09,lirun09l)
	print '''
	本周甲醇制pp现货利润%s，利润%s%s。盘面5月利润%s，利润%s%s。盘面9月利润%s，利润%s%s。
	'''%(lirun,b,a,lirun05,d,c,lirun09,f,e)

	#传统下游
	table = data.sheet_by_name(u'卓创开工率')
	for i in range(table.nrows):
		v=table.cell(i,3)
		a =str(v)
		if "empty" in a and i > 200:
			break
	jiaquanrate = float(str(table.cell(i-1,3))[7:])*100
	jiaquanratel = float(str(table.cell(i-2,3))[7:])*100
	cusuan = float(str(table.cell(i-1,6))[7:])*100
	cusuanl =  float(str(table.cell(i-2,6))[7:])*100
	ejm =  float(str(table.cell(i-1,4))[7:])*100
	ejml = float(str(table.cell(i-2,4))[7:])*100
	a,b=zj(jiaquanrate,jiaquanratel)
	c,d = zj(cusuan,cusuanl)
	e,f=zj(ejm,ejml)
	print """2、传统下游
	甲醛开工负荷%s%%，%s%s%%，
	醋酸开工负荷%s%%，%s%s%%，
	二甲醚开工负荷%s%%，%s%s%%，
	"""%(jiaquanrate,a,b,cusuan,c,d,ejm,e,f)


def main():

	neipan()
	

	waipan()
	
	shui()

	kucun()

	gongji()

	xuqiu()



if __name__ =="__main__":
	main()
