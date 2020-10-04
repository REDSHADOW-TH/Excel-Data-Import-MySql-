//Copyright 2020 REDSHADOWTH(Phumiphat Joreonyonwhatthana) under MIT Licence.
//MIT Licence https://opensource.org/licenses/MIT

using Excel_Data_Import__MySql_.Data;

namespace LuaShare
{
    class ImportData
    {
        public IExcelConfig MakeExcelConfig(string path, string sheet, int startRow)
        {
            return new IExcelConfig()
            {
                Path = path, Sheet = sheet, StartRow = startRow
            };
        }
        public IImportData MakeImportData(string feild, string columnName)
        {
            return new IImportData() { Feild = feild, ColumnName = columnName };
        }
    }
}
