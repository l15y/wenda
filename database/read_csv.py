import csv

def readCSV(filename):   ##读取csv的二维数组，第一行是列名
	List = []
	with open(filename,'r') as csvfile:
		reader = csv.reader(csvfile, delimiter=',')
		for row in reader:
			List.append(row)
	return List
	
def readCSV2(filename):   ##读取csv的二维数组，第一行是列名(空格分割)
	List = []
	with open(filename,'r') as csvfile:
		reader = csv.reader(csvfile, delimiter=' ')
		for row in reader:
			List.append(row)
	return List


def readCSVbyColumn(filename , columnname):   ##读取csv列名对应列，不包括列名
	List = []
	with open(filename,'r') as csvfile:
		reader = csv.reader(csvfile, delimiter=',')
		p = -1
		i = 0
		j = 0
		for row in reader:
			if i == 0:
				for c in row:
					if c == columnname :
						p = j
					j += 1
			else:
				List.append(row[p])
			i += 1
			if p == -1:
				break
	return List
	