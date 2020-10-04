filePath = 'C:\\Users\\nilph\\source\\repos\\Excel Data Import (MySql)\\Excel Data Import (MySql)\\bin\\Debug\\netcoreapp3.1\\File\\Test.xlsx'


excelConfig = {
    path = filePath, sheet = 'Sheet1', startRow = 2
}

data = {
    importData:MakeImportData('id', 'ID'),
    importData:MakeImportData('name', 'Name')
}

system:import(excelConfig, data, 'test1')
