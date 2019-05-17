import re
import xlrd
import xlwt
import copy
	# 读取Excel
import random
import numpy as np
import string
from pymongo import MongoClient
class TestMongo(object):

    def __init__(self):
        self.client = MongoClient()
        self.db = self.client['Biolmex']

    def add_one(self, ID, Interaction_participants, interaction_participants_numble, Mutation_postion,Original_sequence,Resulting_sequence, Mutation_type ):
        # 新增数据
        post = {
			'ID': ID,
			'Interaction_participants':Interaction_participants,
            'interaction_participants_numble':interaction_participants_numble,
            'Mutation_postion':Mutation_postion,
            'Original_sequence':Original_sequence,
            'Resulting_sequence':Resulting_sequence,
            'Mutation_type':Mutation_type,
        }
        return self.db.Biolmex.insert_one(post)


if __name__ == "__main__":
	obj = TestMongo()
	wk = xlrd.open_workbook(r'D:\Bioinformation\Data\data1.xlsx')
	ws = wk.sheet_by_index(0)
	nrows = ws.nrows
	 # 获取总列数
	ncols = ws.ncols
	flag=0
	print(nrows)
	print(ncols)
	for i in range(2, nrows):
		ID = ws.cell(i, 7).value
		Interaction_participants =  ws.cell(i, 11).value
		interaction_participants_numble =  ws.cell(i, 17).value
		Mutation_postion =  ws.cell(i, 2).value
		Original_sequence =  ws.cell(i, 3).value
		Resulting_sequence = ws.cell(i, 4).value
		Mutation_type =  ws.cell(i, 5).value
		rest = obj.add_one(ID, Interaction_participants, interaction_participants_numble, Mutation_postion, Original_sequence,
		Resulting_sequence, Mutation_type)