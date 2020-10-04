-- Copyright 2020 REDSHADOWTH (Phumiphat Joreonyonwhatthana) under MIT Licence.
-- MIT Licence https://opensource.org/licenses/MIT


-- import system std funtion from c# code.
_importData = clr.LuaShare.ImportData
importData = _importData()
_allImportData = nil

system = { excelConfig = nil, importData = nil, tableTarget = nil }

function system:import(config, data, table)
    self.excelConfig = config
    self.importData = data
    self.tableTarget = table
end
