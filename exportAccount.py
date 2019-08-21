import pygsheets
# https://github.com/nithinmurali/pygsheets
class exportAccount:
	def __init__(self):
		self.o = open('import.csv', 'wt')
		gc = pygsheets.authorize(service_file='My First Project-745f05f550a6.json')
		sh = gc.open('生活雜支')
		self.wks = sh.worksheet_by_title("生活雜支2")
		self.pointMatrix = {}
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
			if '簽帳卡' in cell[1].value:
				continue
			self.o.write(cell[0].value + ',')
			self.o.write(cell[1].value + ',')
			self.o.write(cell[2].value + ',')
			self.o.write(cell[5].value + ',')
			self.o.write('{0} {1}'.format(cell[3].value.replace(',',' '),cell[4].value) + ',')
			self.o.write(',\n')
	def close(self):
		self.o.close()	
	def process(self):
		inputData = input("輸入處理區間: ")
		# 寫入標題
		self.o.write(self.getHeadRow() + '\n')

		# 取得雲端生活雜支資料
		items = inputData.split(",")
		for item in items:
			if '-' in item:
				start = item.split("-")[0]
				end = item.split("-")[1]
				cell_list = self.wks.range('A'+start + ':G' + end)
				self.printList(cell_list)
			else:
				cell_list = self.wks.range('A'+item + ':G' + item)
				self.printList(cell_list)
		# 輸出樂天點數帳務
		for key, value in self.pointMatrix.items():
			self.o.write(key + ',')
			self.o.write('下月庫存,')
			self.o.write('其它收入,')
			self.o.write(str(value) + ',')
			self.o.write(key + ' 樂天點數,')
			self.o.write(',\n')
		self.close()

	
if __name__ == '__main__':
	exportAccount().process()


