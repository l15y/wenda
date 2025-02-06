# coding: utf-8
import random

#class word_vector :
#	word = None
#	vector = None
#	
#	def __init__(self, line):
#		line = line.split(" ")
#		self.wv_listword = line[0]
#		self.wv_listvector = []
#		for i in range(1,len(line)):
#			self.wv_listvector.append(line[i])

def cos_simi(vector1,vector2):  
	dot_product = 0.0  
	normA = 0.0  
	normB = 0.0
	
	for a,b in zip(vector1,vector2):  
		dot_product += a*b  
		normA += a**2  
		normB += b**2  
	
		
	if normA == 0.0 or normB==0.0:  
		return None  
	else:  
		return dot_product / ((normA*normB)**0.5)  

class word_vector_model :
	wv = None
	def read_vec(self, vec_src):
		self.wv = {}
		with open(vec_src,'r',encoding="utf-8") as f:
			count = 0
			for line in f.readlines():
				count += 1
				if count % 10000 == 9999:
					print('loading word vector (' + str(count+1) + ') ......')
				line = line.split(" ")
				curList = []
				for i in range(1,len(line)):
					curList.append(float(line[i]))
				self.wv[line[0]] = curList
		print('word vector read over...')
	
	
	def get_simi_top(self, word , top_num):
		top_num += 1
		vec = self.wv[word]
		curword = []
		cursimi = []
		for key,value in self.wv.items():
			if len(key) > 12:
				continue 
			if random.randint(0,100) < 70:  #留百分8数据
				continue 
			if key == word:
				continue
			simi = cos_simi(vec, value)
			curword.append(key)
			cursimi.append(simi)
			continue
			p = len(curword)-1
			if p <= 0:
				curword.append(key)
				cursimi.append(simi)
				continue
				
			while p>=0 :
				if simi < cursimi[p]:
					curword.insert(p+1, key)
					cursimi.insert(p+1, simi)
					break
				if p == 0:
					curword.insert(p, key)
					cursimi.insert(p, simi)
					break
				p-=1
			if len(curword) > top_num:
				curword = curword[:top_num]
				cursimi = cursimi[:top_num]
		
		answord = []
		vis = set()
		top_num -= 1
		for t in range(top_num):
			maxx = -9999999
			maxword = ''
			for i in range(len(curword)):
				if curword[i] in vis:
					continue 
				if cursimi[i]>maxx:
					maxx = cursimi[i]
					maxword = curword[i]
			
			answord.append(maxword)
			vis.add(maxword)
		
		return answord
			
			
			
		 
#wvm = word_vector_model()
#wvm.read_vec('vector.txt')
#print(wvm.get_simi_top('福建', 2))
#print(wvm.get_simi_top('福建省', 2))
#print(wvm.get_simi_top('柳树', 2))