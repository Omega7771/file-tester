import subprocess
import os
import xlsxwriter
import time
import zipfile
import easygui
import sys
from colorama import *
init()
# -*- coding: utf-8 -*-
os.system("title File tester")
def check(name):
	m=''
	f=open(name)
	for line in f:
		if line.find("system")!=-1 and line.find("//")!=0:
			m = m + "//" + line 
			print("!!!System command found... Fix it!!!")
		else:
			m = m + line 
	f.close()
	f=open(name, "w")
	f.write(m)
	f.close()
def inpt(numd):
	global ind
	f=open(f"{fl}input.txt")
	lines=f.readlines()
	f.close()
	f=open("input.txt", "w")
	for i in range(0,tet):
		f.write(lines[ind])
		ind+=1
def start(tests):
	if not os.path.exists(f"{fl}input.txt"):
		print(Fore.RED + "File input.txt not found")
		print(Fore.WHITE + "", end="")
		os.system("pause")
		sys.exit(0)
	f=open(f"{fl}input.txt")
	a=f.readlines()
	f.close()
	if len(a)/tet<tests:
		print(Fore.RED + "Not enough data in the input.txt")
		print(Fore.WHITE + "", end="")
		os.system("pause")
		sys.exit(0)
	for i in range(1,tests+1):
		if not os.path.exists(f"{fl}output({i}).txt"):
			print(Fore.RED + "Not enough outputs file")
			print(Fore.WHITE + "", end="")
			os.system("pause")
			sys.exit(0)
print(Fore.GREEN + "Select directory with cpp files")
print(Fore.WHITE + "", end="")
os.system("pause")
os.chdir(easygui.diropenbox())
print(Fore.GREEN + "Select directory with input and output files")
print(Fore.WHITE + "", end="")
os.system("pause")
fl=easygui.diropenbox()+"\\"
numb=int(input("Enter the number of tests: "))
start(numb)
tet=int(input("Enter the number of lines for one input: "))
c=os.listdir()
ind=0
start_time = time.time()
i=0
name=""
nam=""
k=0
res=0
m=8.43
while True:
	if i==len(c): break
	if str(c[i]).find(".cpp")==-1:
		del c[i]
	else: i=i+1
k=i
workbook = xlsxwriter.Workbook('Results.xlsx')
bas = workbook.add_format()
bas.set_align('center')
acc = workbook.add_format()
acc.set_bold()
acc.set_align('center')
acc.set_font_color('green')
err = workbook.add_format()
err.set_bold()
err.set_align('center')
err.set_font_color('red')
som = workbook.add_format()
som.set_bold()
som.set_align('center')
som.set_font_color('orange')
ex = workbook.add_worksheet()
ex.set_column(1, numb, 17)
ex.write(0, 0, 'Name', bas)
for i in range(0,numb):
	ex.write(0, i+1, f'Test {i+1}', bas)
ex.write(0, numb+2, "Result", bas)
for g in range(0,len(c)):
	ind=0
	res=0
	name=c[g]
	nam=name[0:len(name)-4]
	print(f"=======Testing {name}=======")
	check(name)
	if len(nam)>m:
		m=len(nam)
		ex.set_column(0, 0, m)
	ex.write(g+1, 0, nam, bas)
	try:
		subprocess.call(f"g++ \"{name}\" -o \"{nam}.exe\"")
		if os.path.exists(f"{nam}.exe"):
			for i in range(0,numb):
				inpt(tet)
				print(f"------Test {i+1}------")
				f=open(f"input.txt")
				stu=f.read()
				if stu[len(stu)-1]=="\n":
					stu=stu[0:len(stu)-1]
				print(f"Input: {stu}")
				f.close()
				try:
					prog=subprocess.Popen(f"\"{nam}.exe\" < input.txt > \"output_{nam}_test_{i+1}.txt\"",shell=True).wait(timeout=1)
					f=open(f"output_{nam}_test_{i+1}.txt")
					a=f.read()
					r=open(f"{fl}output({i+1}).txt")
					print(f"Output: {a}")
					if r.read()==a:
						ex.write(g+1, i+1, "Accepted", acc)
						res=res+1
					else:
						ex.write(g+1, i+1, "Wrong answer", err)
					f.close()
					r.close()
				except Exception as e:
					a=str(e)
					if a.find("timed out")!=-1:
						os.system(f"taskkill /im \"{nam}.exe\" /f")
						ex.write(g+1, i+1, "timeout", err)
						print("timeout")
					else:
						f=open(f"output_{nam}_test_{i+1}.txt","w")
						f.write(str(e))
						f.close()
		else:
			f=open(f"output_{nam}_test_{i+1}.txt","w")
			f.write("Compilation error")
			f.close()
			for i in range(0,numb):
				ex.write(g+1, i+1, "Compilation error", err)
		rus=round((res/numb)*100, 1)
		if rus==100:
			ex.write(g+1, numb+2, f"{int(rus)}%", acc)
		elif rus!=0:
			ex.write(g+1, numb+2, f"{int(rus)}%", som)
		else:
			ex.write(g+1, numb+2, f"{int(rus)}%", err)
	except Exception as e:
		f=open(f"output_{nam}_test_{i+1}.txt","w")
		f.write("Compilation error")
		f.close()
	z = zipfile.ZipFile(f'{nam}.zip', 'w')
	z.write(f'{nam}.cpp')
	for i in range(0,numb):
		if os.path.exists(f'output_{nam}_test_{i+1}.txt'):
			z.write(f'output_{nam}_test_{i+1}.txt')
			os.remove((f'output_{nam}_test_{i+1}.txt'))
	z.close()
	if os.path.exists(f"{nam}.exe"): os.remove(f"{nam}.exe")
workbook.close()
z = zipfile.ZipFile(f'output.zip', 'w')
for g in range(0,k):
	nam=c[g][0:len(c[g])-4]
	z.write(f'{nam}.zip')
	os.remove(f'{nam}.zip')
z.close()
print("==============")
time=(time.time() - start_time)
print(f"Tested {len(c)} file on {numb} tests in {int(time/60)}:{int(time%60)} minutes ")
os.remove("input.txt")
d=input("Do you want open results?(Y/N): ")
if d=="Y" or d=="y":
	os.startfile("Results.xlsx")
os.system("pause")
