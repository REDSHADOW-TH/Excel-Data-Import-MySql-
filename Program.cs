//Copyright 2020 REDSHADOWTH(Phumiphat Joreonyonwhatthana) under MIT Licence.
//MIT Licence https://opensource.org/licenses/MIT

using System;
using Excel_Data_Import__MySql_.Data;
using Excel_Data_Import__MySql_.Lua;
using Excel_Data_Import__MySql_.Excel;


namespace Excel_Data_Import__MySql_
{
    class Program
    {
        static void Main(string[] args)
        {
            DataStore._config = Config.GetConfig();
            LuaProvider luaProvider = new LuaProvider();

            IExcelConfig excelConfig = DataStore._excelConfig;
            // check excel config.
            if (!String.IsNullOrEmpty(excelConfig.Path) || !String.IsNullOrEmpty(excelConfig.Sheet) || excelConfig.StartRow != null)
            {
                // check import data.
                if (DataStore._importData != null && DataStore._importData.Count > 0)
                {
                    // check table target.
                    if (!String.IsNullOrEmpty(DataStore._tableTarget))
                    {
                        Console.WriteLine("Data ready for import.");
                        Import excelImport = new Import();
                        excelImport.StartImport();
                    } else
                    {
                        Console.WriteLine("Table target is not null or empty.");
                    }
                } else
                {
                    Console.WriteLine("Import data is empty");
                }
            } else
            {
                Console.WriteLine("Excel config is not valid (path, sheet, startRow is not null or empty).");
                Console.ReadKey();
                Environment.Exit(1);
            }

            Console.ReadKey();
        }
    }
}
