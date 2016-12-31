import win32com.client as win32

# Opens xls, deletes all irrelevant columns in ALL (22) sheets, saves as xlsx.

excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = False

for i in range(2004,2016):

	wb = excel.Workbooks.Open(r'C:\Users\Konstantinos\Documents\GIT\Python Projects\Goal Supremacy Data Testing\data\all-euro-data-%d-%d.xls' % (i, i+1) )	# Add own path and files

	to_keep = ['Div', 'Date', 'HomeTeam', 'AwayTeam', 'FTHG', 'FTAG', 'B365H', 'B365D', 'B365A']															# Columns to keep

	for n in range(22) :

		ws = wb.Sheets(n+1)

		counter = 1

		for r in range(75):

			print(ws.Cells(1, counter).Value, counter)

			if ws.Cells(1, counter).Value not in to_keep:

				print(ws.Cells(1, counter).Value, counter, "DEL")

				ws.Columns(counter).Delete()

			else:
				counter += 1




	wb.SaveAs(r'C:\Users\Konstantinos\Documents\GIT\Python Projects\Goal Supremacy Data Testing\data\all-euro-data-%d-%d.xlsx' % (i, i+1), FileFormat=win32.constants.xlOpenXMLWorkbook)		# FileFormat is needed to convert file to .xlsx - Add own path and filename
	excel.Application.Quit()
