import json
import xlsxwriter

with open('ClueBackup-2017-08-19.cluedata', 'r') as clue_backup:
	json_structure = json.loads(clue_backup.read())

	known_columns = list()

	# collect columns
	for day in json_structure['data']:
		for key in day.keys():
			if key not in known_columns:
				known_columns.append(key)
	
	# print data
	
	workbook = xlsxwriter.Workbook('cluedata.xlsx')
	worksheet = workbook.add_worksheet()

	col = 0
	for column in known_columns:
		worksheet.write(0, col, column)
		col += 1

	row = 1
	for day in json_structure['data']:
		for key in day.keys():
			if isinstance(day[key], list):
				worksheet.write(row, known_columns.index(key), ', '.join(day[key]))
			elif isinstance(day[key], dict):
				worksheet.write(row, known_columns.index(key), str(day[key]))
			else:
				worksheet.write(row, known_columns.index(key), day[key])
		row += 1
	
	worksheet.freeze_panes(1, 1)
	workbook.close()
	
