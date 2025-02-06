import sys
import os

import pymongo
from pymongo import MongoClient
from plugins.common import settings
class Mongo():
	clent = None
	db = None
	collection = None
	def makeConnection(self):
		cunnrent_setting = settings.librarys.kg
		mongo_host = cunnrent_setting.mongo_host
		self.client = MongoClient(mongo_host,27017)

	def getDatabase(self,dbName):
		self.db = self.client[dbName]
		return self.db

	def getCollection(self,collectionName):
		self.collection = self.db[collectionName]
		return self.collection

