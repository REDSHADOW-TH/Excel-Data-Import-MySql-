//Copyright 2020 REDSHADOWTH(Phumiphat Joreonyonwhatthana) under MIT Licence.
//MIT Licence https://opensource.org/licenses/MIT

using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using OfficeOpenXml;
using Excel_Data_Import__MySql_.Data;
using Excel_Data_Import__MySql_.MySql;
using System.Linq;

namespace Excel_Data_Import__MySql_.Excel
{
    class Import
    {
        private string currentDir = Directory.GetCurrentDirectory();

        private string exception = "";

        // data from data store.
        private IExcelConfig excelConfig = DataStore._excelConfig;
        private List<IImportData> importData = DataStore._importData;
        private string tableTarget = DataStore._tableTarget;

        // declare variable for epplus.
        private ExcelPackage _excelInstance;
        private ExcelWorkbook _workbook;
        private ExcelWorksheet _sheetTarget;

        private Dictionary<string, int> feildMapping = new Dictionary<string, int>();
        public Import()
        {
            if (File.Exists(excelConfig.Path))
            {
                _excelInstance = new ExcelPackage(new FileInfo(@excelConfig.Path));
                _workbook = _excelInstance.Workbook;
                _sheetTarget = _workbook.Worksheets[@excelConfig.Sheet];

                if (_sheetTarget == null)
                {
                    Console.WriteLine($"Can not find sheet {excelConfig.Sheet} please check config and exel file.");
                    Console.ReadKey();
                    Environment.Exit(1);
                }
            } else
            {
                Console.WriteLine($"Can not find path {excelConfig.Path}");
                Console.ReadKey();
                Environment.Exit(1);
            }
         
            GetFeildMapping();
        }

        public bool StartImport()
        {
            List<string> allFeildsData = importData.Select(item => item.Feild).Distinct().ToList();
            string allFeilds = MakeField();
            string prefixInsert = $"insert into {tableTarget} values ";

            bool firstValue = true;


            string sqlInsertCommand = "";

            StringBuilder allValues = new StringBuilder();
            try
            {
                int endRow = _sheetTarget.Dimension.End.Row;

                // start loop from excel config (startRow)
                for(int row=(int)excelConfig.StartRow;row<=endRow;row++)
                {

                    if (firstValue)
                    {
                        allValues.AppendLine(MakeValue(row));
                        firstValue = false;
                    } else
                    {
                        allValues.AppendLine("," + MakeValue(row));
                    }
                }
                sqlInsertCommand = prefixInsert + allValues.ToString();

                IMySqlCommand mysqlExec = new IMySqlCommand() { Command = sqlInsertCommand };

                if (mysqlExec.Execute())
                {
                    Console.WriteLine("Import success.");
                    Console.ReadKey();
                    Environment.Exit(0);
                } else
                {
                    Console.WriteLine($"Import fail exception: {mysqlExec.GetException()}");
                    Console.ReadKey();
                    Environment.Exit(1);
                }
                return true;
            } catch(Exception ex)
            {
                exception = ex.ToString();
                return false;
            }

            string MakeField()
            {
                string result = "";
                bool firstValue = true;
                if (allFeildsData.Count > 0)
                {
                    foreach(string item in allFeildsData)
                    {
                        if (firstValue)
                        {
                            result += item;
                            firstValue = false;
                        } else
                        {
                            result += "," + item;
                        }
                    }
                }
                return result;
            }

            string MakeValue(int row)
            {
                string result = "(";

                bool firstValue = true;
                foreach (string item in allFeildsData) 
                {
                    // get column name by field.
                    string key = importData.Where(w => w.Feild == item).Select(s => s.ColumnName).FirstOrDefault().ToString();

                    int column = 0;
                    try
                    {
                        // get column index from mmapping.
                         column = feildMapping[key];
                    } catch
                    {
                        Console.WriteLine($"Not found column {key} please check import data and try again.");
                        Console.ReadKey();
                        Environment.Exit(1);
                    }
                    // get value from excel.
                    dynamic value = _sheetTarget.Cells[row, column].Value != null ? _sheetTarget.Cells[row, column].Value.ToString() : "";

                    if (firstValue)
                    {
                        result += $"'{value}'";
                        firstValue = false;
                    } else
                    {
                        result += $",'{value}'";
                    }
                }

                    result += ")";
                return result;
            }
        }

        private void GetFeildMapping()
        {
            Dictionary<string, int> result = new Dictionary<string, int>();

            int maxCol = _sheetTarget.Dimension.End.Column;

            for(int index = 1; index <= maxCol; index++)
            {
                // key = value in cell, value = index (column index).
                string key = _sheetTarget.Cells[1, index].Value.ToString();
                feildMapping.Add(key, index);
            }
        }
    }
}
