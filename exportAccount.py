import pygsheets
# https://github.com/nithinmurali/pygsheets
class exportAccount:
	def __init__(self):
		self.o = open('import.csv', 'wt')
		gc = pygsheets.authorize(service_file='My First Project-745f05f550a6.json')
		sh = gc.open('生活雜支')
		self.wks = sh.worksheet_by_title("生活雜支2")
		self.pointMatrix = {}
		self.cell_list = []
	def filter(self):
		less_list = []
		for cell in self.cell_list:
			if cell[0].value != '' and cell[7].value == '' :
				less_list.append(cell)
			elif cell[0].value  == '':
				break
		return less_list
	def getHeadRow(self):
		return '記錄日期,目的項目,來源項目,異動金額,記錄摘要,詳細備註,發票號碼'
	def printList(self,lst):
		for cell in lst:
			# 累計樂天點數			
			if cell[6].value != '':
				point = int(cell[6].value)
				date = cell[0].value
				if date in self.pointMatrix:
					self.pointMatrix[date] = self.pointMatrix[date] + point
				else:
					self.pointMatrix[date] = point
			# 輸出帳務資料	
			self.o.write(cell[0].value + ',')
			self.o.write(cell[1].value + ',')
			self.o.write(cell[2].value + ',')
			self.o.write(cell[5].value.replace(',','') + ',')
			self.o.write('{0} {1}'.format(cell[3].value,cell[4].value) + ',')
			self.o.write(',\n')			
	def close(self):
		self.o.close()	
	def process(self):
		# 寫入標題
		self.o.write(self.getHeadRow() + '\n')

		# 取得雲端生活雜支資料
		self.cell_list = self.wks.range('A1900:H2500')
		self.printList(self.filter())
		 
		# 輸出樂天點數帳務
		for key, value in self.pointMatrix.items():
			self.o.write(key + ',')
			self.o.write('下月庫存,')
			self.o.write('網拍收入,')
			self.o.write(str(value) + ',')
			self.o.write(key + ' 樂天點數,')
			self.o.write(',\n')
		self.close()

	
if __name__ == '__main__':
	exportAccount().process()


