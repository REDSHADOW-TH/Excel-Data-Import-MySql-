//Copyright 2020 REDSHADOWTH(Phumiphat Joreonyonwhatthana) under MIT Licence.
//MIT Licence https://opensource.org/licenses/MIT

using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using Excel_Data_Import__MySql_.Data;
using ironLua = Neo.IronLua;

namespace Excel_Data_Import__MySql_.Lua
{
    class LuaProvider
    {
        public static string _code;
        public static dynamic _env = null;


        private string currentPath = Directory.GetCurrentDirectory();
        private string scriptPath;

        public LuaProvider()
        {
            if (File.Exists(DataStore._config.ScriptPath))
            {
                if (!String.IsNullOrEmpty(DataStore._config.ScriptPath))
                {
                    scriptPath = DataStore._config.ScriptPath;
                }
                else
                {
                    Console.WriteLine("Config is not valid scriptPath is not null or empty");
                    Console.ReadKey();
                    Environment.Exit(1);
                }
            } else
            {
                Console.WriteLine($"Can not find path {DataStore._config.ScriptPath}");
                Console.ReadKey();
                Environment.Exit(1);
            }
            OnStart();
        }

        public void OnStart()
        {
            string code = "";
            using (StreamReader file = new StreamReader(scriptPath))
            {
                code = file.ReadToEnd();
            }


            StringBuilder stdCode = new StringBuilder();

            // std code.
            stdCode.AppendLine("_importData = clr.LuaShare.ImportData");
            stdCode.AppendLine("importData = _importData()");
            stdCode.AppendLine("_allImportData = nil");
            stdCode.AppendLine("system = { excelConfig = nil, importData = nil, tableTarget = nil }");
            stdCode.AppendLine("function system:import(config, data, table)\n" +
                    "self.excelConfig = config\n" +
                    "self.importData = data\n" +
                    "self.tableTarget = table\n" +
                "end");


            if (!String.IsNullOrEmpty(code))
            {
                stdCode.AppendLine(code);
            }

            _code = stdCode.ToString();

            ironLua.Lua lua = new ironLua.Lua();
            dynamic env = lua.CreateEnvironment();

            env.dochunk(_code);
            if (env != null)
            {
                _env = env;
            }


            // get excel config.
            if (env.system.excelConfig != null)
            {
                dynamic config = env.system.excelConfig;
                IExcelConfig data = new IExcelConfig();
                bool pathCheck, sheetCheck, startRowCheck, EndRowIgnoreCheck;
                pathCheck = sheetCheck = startRowCheck = EndRowIgnoreCheck = false;
                foreach (KeyValuePair<dynamic, dynamic> item in config)
                {
                    dynamic key = item.Key;
                    dynamic value = item.Value;
                    if (key == "path" && value != null)
                    {
                        data.Path = value;
                        pathCheck = true;
                    }
                    else if (key == "sheet" && value != null)
                    {
                        data.Sheet = value;
                        sheetCheck = true;
                    }
                    else if (key == "startRow" && value != null)
                    {
                        data.StartRow = (int)value;
                        startRowCheck = true;
                    }
                    else if (key == "endRowIgnore" && value != null)
                    {
                        data.EndRowIgnore = (int)value;
                        EndRowIgnoreCheck = true;
                    }
                }

                if (!EndRowIgnoreCheck)
                {
                    data.EndRowIgnore = 0;   
                }

                if (!pathCheck || !sheetCheck || !startRowCheck)
                {
                    Console.WriteLine("Excel config is not valid (path, sheet, startRow is not null or empty).");
                    Console.ReadKey();
                    Environment.Exit(1);
                } else
                {
                    DataStore._excelConfig = data;
                }
            }



            // get import data.
            if (env.system.importData != null)
            {
                List<IImportData> data = new List<IImportData>();
                foreach (KeyValuePair<dynamic, dynamic> item in env.system.importData)
                {
                    if (!String.IsNullOrEmpty(item.Value.Feild) && !String.IsNullOrEmpty(item.Value.ColumnName))
                    {
                        data.Add(new IImportData()
                        {
                            Feild = item.Value.Feild,
                            ColumnName = item.Value.ColumnName
                        });
                    }
                }
                if (data.Count > 0)
                {
                    DataStore._importData = data;
                }

            }

            // get table target.
            if (env.system.tableTarget != null)
            {
                string tableTarget = (string)env.system.tableTarget;
                if (String.IsNullOrEmpty(tableTarget))
                {
                    Console.WriteLine("Table target is not null or empty.");
                } else
                {
                    DataStore._tableTarget = tableTarget;
                }
            }
        }
    }
}
