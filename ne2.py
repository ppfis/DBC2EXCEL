#!/usr/bin/env python2.7
#-*-coding:utf-8-*-
#__author__penghuayuan
#__version__=0.1
#__date__2017/08/08

#-*-本程序用于将特定形式的EXCEL（对EXCEL要求）转换为DBC格式-*-
#已知BUG#1.signal_value_description中不允许有中文字符’：‘，必须在转换前替换成英文字符；--会报错，无法生成；
		#2.不能处理用’;'分隔不同signal_value的情形，只支持回车换行的形式；--对应部分的DBC有问题，不影响生成；

import os
import xlrd
def read_xls():
	#首行名称的定义，用于区别哪行文字是什么
	msg_name='Msg Name'
	msg_type='Msg Type'
	msg_id='Msg ID'
	msg_send_type='Msg Send Type'
	msg_cycle='Msg Cycle'
	msg_len='Msg Length'
	signal_name='Signal Name'
	signal_description='Signal Description'
	msg_cycle='Msg Cycle'
	byte_oder='Byte Order'
	start_byte='Start Byte'
	start_bit='Start Bit'
	signal_send_type='Signal Send Type'
	bit_len='Bit Length'
	data_type='Data Type'
	resolution='Resolution'
	signal_max='Signal Max. Value (phys)'
	signal_min='Signal Min. Value (phys)'
	signal_value_description='Signal Value Description'
	offset='Offset'
	unit='Unit'
	#EXCEL文件的读取
	workbook=xlrd.open_workbook(r'BCAN.xls')
	sheet5=workbook.sheet_by_name('Matrix')
	row=sheet5.nrows
	col=sheet5.ncols
	i=0
	j=0
	title=sheet5.row_values(0)#把第一行读取出来
	#查找有多少个节点，这里规定节点必须从29行开始，不能是29行之前
	MY_NODE=[]
	for m in title[29:]:
		MY_NODE.append(m)
		MY_NODE.append(' ')
		
		
	
	with open('my_bcan.dbc','w') as f:
		signal_min1=[]
		signal_max1=[]
		signal_description1=[]
		signal_value_description1=[]
		node=[]
		signal_sg=[]
		node_num=0
		signal_cmd=[]
		dict={}
		f.write('VERSION \"\"\n\n')
		f.write('NS_ :\n'+8*' '+'CM_'+'\n\n'+'BS_:'+'\n\n')
		f.write('BU_: ')
		f.writelines(MY_NODE)
		f.write('\n\n')
		
		
		for j in range(col):
			l=sheet5.col_values(j)
			#print type(l[0])
			l[0]=l[0].encode('utf-8')
			if msg_name in l[0]:
				msg_name1=l
			elif msg_id in l[0]:
				msg_id1=l
			elif msg_len in l[0]:
				msg_len1=l
			elif signal_name in l[0]:
				signal_name1=l
			elif signal_description in l[0]:
				signal_description1=l
			elif start_bit in l[0]:
				start_bit1=l
			elif bit_len in l[0]:
				bit_len1=l
			elif byte_oder in l[0]:
				byte_oder1=l
			elif start_byte in l[0]:
				start_len1=l
			elif signal_send_type in l[0]:
				signal_send_type1=l
			elif data_type in l[0]:
				data_type1=l
			elif resolution in l[0]:
				resolution1=l
			elif signal_value_description in l[0]:
				
				signal_value_description1=l
				#print signal_value_description1[3]
			elif offset in l[0]:
				offset1=l
			elif signal_max in l[0]:
				signal_max1=l
			elif signal_min in l[0]:
				signal_min1=l
			elif unit in l[0]:
				unit1=l
			else:
				node.append(l)
				#print (node)
				node_num+=1
		
			
		x=1
		fir=1
		
		NODE=['ok']
		SIGNAL=[['OK']]
		val=[]
		while x<row:
			y=0
			count=0
			NODE.append('ok')
			SIGNAL.append(['OK'])
			sg_inner=''
			sg_header=''
			sg_tail=''
			signal_val=''
			
			if msg_name1[x]=='':
				if signal_name1[x]!='':
					signal_sg.append(msg_id_now+' '+signal_name1[x].encode('utf-8').strip()+' '+'\"'+signal_description1[x].replace('\"',' '))
					sg_header=' SG_ '+signal_name1[x].encode('utf-8').strip()+' : '+str(int(start_bit1[x]))+'|'+str(int(bit_len1[x]))+'@'+'0'+'+ '
					sg_inner='('+str(resolution1[x])+','+str((offset1[x]))+') '
					sg_tail='['+str(signal_min1[x])+'|'+str(signal_max1[x])+']'+' \"'+unit1[x].encode('utf-8')+'\"'+' '
					#print signal_min1[x]'
					f.write(sg_header+sg_inner+sg_tail)
					#信号值描述的说明：
					if signal_value_description1[x].strip()!='' and signal_value_description1[x].count(':')!=0 and signal_value_description1[x].count('0x')!=0:#val用来放信号值的描述signal_value_description。把‘“’，中文的冒号，替换掉
						signal_val=signal_value_description1[x].encode('utf-8')
						signal_val=signal_val.replace('\"',' ')
						signal_val=signal_val.replace('：',':')
						#print(signal_val),
						
						val.append(msg_id_now+' '+signal_name1[x].encode('utf-8').strip()+' '+find_val(signal_val).encode('utf-8'))
					while y<node_num:
						if node[y][x]=='r'or node[y][x]=='R':
							
							SIGNAL[x][count]=(str(node[y][0]))
							#print SIGNAL[x][:count-1]
							SIGNAL[x].append(',')
							SIGNAL[x].append('OK')
							count+=2
						
							
						y+=1
					
						
					
					f.writelines(SIGNAL[x][:count-1])
					if count==0:
						f.write('Vector__XXX')
					f.write('\n')
				else:
					pass
				
			else:
				
				y=0
				
				while y<node_num:
					if node[y][x]=='s'or node[y][x]=='S':
						NODE[x]=(str(node[y][0]))
					y+=1
					msg_id_now=str(int(msg_id1[x],16))
				f.write('\n'+'BO_'+' '+str(int(msg_id1[x],16))+' '+msg_name1[x].encode('utf-8')+': '+str(int(msg_len1[x]))+' '+NODE[x]+'\n')	
				if signal_description1[x]!='':
					signal_cmd.append(str(int(msg_id1[x],16))+' \"'+signal_description1[x])
			
					
			x+=1
		#CM_ "Version:V03 , Date: , Author: , Review: , Approval: , Description: ; ";
		f.write('\nCM_ '+'\"'+"Version:V03 , Date: , Author: , Review: , Approval: , Description: ;"+'\"'+';'+'\n')
		for signal_cm in signal_cmd:
			f.write('CM_ BO_ '+signal_cm+'\"'+';'+'\n')
		#for signal_s in signal_sg:
		k=0
		while k<len(signal_sg):
			
			f.write('CM_ SG_ '+signal_sg[k].encode('utf-8')+'\"'+';'+'\n')
			k+=1
		for v in val:
			f.write('VAL_ '+v.encode('utf-8')+';\n')

	i+=1		
def find_val(s):#此函数用来处理signal_value_description,在list中找出signal_value和signal_description
	head='0x'
	inner1=':'
	inner2='~'
	tail='\n'
	oth=';'
	i=0
	k=0
	d=0
	j=0
	c=0
	s_val=''
	if s.count(inner2)==0 and s.count(inner1)!=0:#处理不含‘~’的部分，把数字提取出来
		while i<s.count(inner1):
		#一部分’;‘结尾而不用换行符'\n'来区分的行为
			val_num=str(int(s[s.find(head,j):s.find(inner1,j)],16))
			val_string=s[(s.find(inner1,j)+1):s.find(tail,j)].strip()
			s_val=val_num+'\"'+val_string+'\"'+s_val
			j=s.find(tail,j)+1
			i+=1
			
	elif s.count(inner2)!=0 and s.count(inner1)!=0:#处理不含‘~’的部分，还要预防~不止一个的情形,还要预防，‘~’在中间的情形
		val=''
		j=0
		c=s.count(inner2)
		
		while d<s.count(inner1):
			if s.find(inner1,j)<s.find(inner2,j) or s.count(inner2,j)==0:#因为~个数很可能先变成0，所以要考虑周全；
				val_num=str(int(s[s.find(head,j):s.find(inner1,j)],16))
				val_string=s[(s.find(inner1,j)+1):s.find(tail,j)].strip()
				s_val=val_num+'\"'+val_string+'\"'+s_val
				j=s.find(tail,j)+1
			elif s.find(inner1,j)>s.find(inner2,j) and s.count(inner2,j)!=0:
				while k<int(s[s.find(inner2,j)+1:s.find(inner1,j)],16)-int(s[s.find(head,j):s.find(inner2,j)],16)+1:
					val_num=str(int(s[s.find(head,j):s.find(inner2,j)],16)+k)
					val_string=s[(s.find(inner1,j)+1):s.find(tail,j)].strip()
					s_val=val_num+'\"'+val_string+'\"'+s_val
					k+=1
					c-=1
				j=s.find(tail,j)+1
				print(s[s.find(tail,j)+1:s.find(tail,j)+16])
			d+=1
	else:
		pass
	return s_val


if __name__=='__main__':
	read_xls()