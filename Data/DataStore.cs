//Copyright 2020 REDSHADOWTH(Phumiphat Joreonyonwhatthana) under MIT Licence.
//MIT Licence https://opensource.org/licenses/MIT

using System.Collections.Generic;

namespace Excel_Data_Import__MySql_.Data
{
    public static class DataStore
    {
        public static IExcelConfig _excelConfig; 
        public static List<IImportData> _importData;
        public static string _tableTarget;
        public static IConfig _config;
    }
}
