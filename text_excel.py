import xlwt
import xlrd
import os

book = xlwt.Workbook()
style0 = xlwt.easyxf('font: name Verdana')
#path = '/suvendu/python/extract_proxy/input/'
print("Enter path of the proxy configuration files :")
path = input()
print("Enter url pattern you want to extract :")
pattern_str = input()

arr_files = [x for x in os.listdir(path) if x.endswith(".conf")]

for findex in range(len(arr_files)):
	ws = book.add_sheet(arr_files[findex])
	print(path + arr_files[findex])
	f = open(path + arr_files[findex])
	data = f.readlines()
	rowcount = 0
	for i in range(len(data)):
		if data[i].strip().startswith('#') or pattern_str not in data[i] or len(data[i]) < 2:
			continue
		row = data[i].strip().split()
		if(len(row) < 3):
			continue
		for j in range(len(row)):
			ws.write(rowcount, j, row[j], style0)
		rowcount += 1
outputfile = path + '/location' + '.xls'
book.save(outputfile);
print("Success!! Output : %s" %(outputfile))
f.close()