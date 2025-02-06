import random

class TREE :
	edge = None  # 层次树邻接表
	leaf = None  # 记录叶子节点
	root = '农业'  # 树的根结点
	curpath = None  # 路径栈
	anspath = None  # 返回路径结果
	UI_str = None # 动态生成的树状图代码
	
	def read_edge(self, src):
		self.edge = {}  # 层次树邻接表
		vis = set()  # 去重
		with open(src,'r',encoding="utf-8") as f:
			for line in f.readlines():
				if line in vis:
					continue
				vis.add(line)
				cur = line.strip().split(' ')
				u = str(cur[0])
				v = str(cur[1])
				if u not in self.edge:
					self.edge[u] = []
				self.edge[u].append(v)
				
	def read_leaf(self, src):
		self.leaf = {}  # 记录叶子节点
		vis = set()   # 去重
		with open(src,'r',encoding="utf-8") as f:
			for line in f.readlines():
				if line in vis:
					continue
				vis.add(line)
				cur = line.strip().split(' ')
				u = str(cur[0])
				v = str(cur[1])
				if u not in self.leaf:
					self.leaf[u] = []
				self.leaf[u].append(v)
			
	def DFS(self, word, u):
		#print(u)
		self.curpath.append(u)
		if u not in self.leaf or word not in self.leaf[u]: # 如果叶节点儿子中没有目标
			pass
		else:  #返回路径
			path = []
			for p in self.curpath:
				path.append(p)
			path.append(word)
			self.anspath.append(path)  # 把该路径加入答案
			
		if u not in self.edge: #儿子只有叶节点了
			pass
		else: # 递归儿子
			for v in self.edge[u]:
				self.DFS(word, v)
		self.curpath.pop()
										#返回根到叶子节点的路径
	def get_path(self, word, unique):  # 可能存在多条路径，所以返回二维数组[路径数][路径]
		self.anspath = []            #unique 为true 代表筛选路径，去除过多重复的路径
		self.curpath = []
		self.DFS(word, self.root)
		random.shuffle(self.anspath)
		if unique == True :
			for i in range(len(self.anspath)):
				j = i + 1
				while j < len(self.anspath):
					seti = set(self.anspath[i])
					setj = set(self.anspath[j])
					unum = len(seti & setj)
					if unum > 2:
						del self.anspath[j]
					else:
						j += 1
		return self.anspath
		
	def get_father(self, word):   # 获得word结点的所有父节点
		ansList = []
		for k,v in self.edge.items(): 
			if word in v:
				ansList.append(k)
		return ansList
		
	def get_branch(self, word):  # 获得word结点的非叶儿子结点
		ansList = []
		for k,v in self.edge.items():
			if k == word:
				ansList = v
				break
		return ansList
		
	def get_leaf(self, word):  # 获得word结点的所有叶子儿子
		if word not in self.leaf:
			return []
		return self.leaf[word]
		
	def DFS_create_UI(self, u, depth):  # 通过dfs生成树形结构代码
		cur = ''
		if len(self.anspath) > depth and u == self.anspath[depth]:
			self.UI_str += ' <li> <span>'
			cur = '<i class="fa fa-minus-square" aria-hidden="true"></i>&nbsp;'
			if len(self.anspath) == depth + 1:
				cur = '<i class="fa fa-plus-square" aria-hidden="true"></i>&nbsp;'
		else:
			self.UI_str += ' <li style="display: none;"> <span>'
			cur = '<i class="fa fa-plus-square" aria-hidden="true"></i>&nbsp;'
		if u in self.edge and len(self.edge[u]) > 0:
			self.UI_str += cur
		else:
			pass
		self.UI_str += str(u) + '</span>' 
		if str(u) == self.anspath[len(self.anspath)-1]:
			self.UI_str += '&nbsp;&nbsp;&nbsp;当前分类'
		else:
			self.UI_str += '&nbsp;<a href="overview?node=' + str(u) + '">&nbsp;&nbsp;[进入分类]</a>'
		
		if u in self.edge and len(self.edge[u]) > 0:
			self.UI_str += '<ul>'
			for v in self.edge[u]:
				self.DFS_create_UI(v, depth+1)
			self.UI_str += '</ul>'
				
		self.UI_str += '</li>'
		
	def DFS2(self, theme, u):
		n = len(self.anspath)
		if n>0 and self.anspath[n-1] == theme:
			return
		self.anspath.append(u)
		if u in self.edge:
			for v in self.edge[u]:
				self.DFS2(theme, v)
		n = len(self.anspath)
		if n>0 and self.anspath[n-1] == theme:
			return
		self.anspath.pop()
		
	def create_UI(self,theme):  # 生成UI代码
		self.anspath = []
		self.DFS2(theme, self.root)  # 此时self.anspath为目标路径(一条)
		self.UI_str = "<ul>"
		self.DFS_create_UI(self.root,0)
		self.UI_str += "</ul>"
		return self.UI_str
		
		
# 读取农业层次树
#tree = TREE()
#tree.read_edge('micropedia_tree.txt')
#tree.read_leaf('leaf_list.txt')
#print(tree.get_path('香蕉',True))

		
		
		
				