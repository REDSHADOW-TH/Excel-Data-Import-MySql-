//Copyright 2020 REDSHADOWTH(Phumiphat Joreonyonwhatthana) under MIT Licence.
//MIT Licence https://opensource.org/licenses/MIT

using System;
using Excel_Data_Import__MySql_.Data;
using MySql.Data.MySqlClient;

namespace Excel_Data_Import__MySql_.MySql
{
    public struct IMySqlCommand
    {
        public string Command;
    }
    public static class MySqlExtension
    {
        private static string exception = "";

        private static MySqlConnection connection = new MySqlConnection(DataStore._config.ConnectionString);
        public static bool Execute(this IMySqlCommand commandData)
        {
            try
            {
                connection.Open();
                char singleQuote = (char)39;
                string insertCommand = commandData.Command.Replace("'", @$"\{singleQuote}");
                MySqlCommand cmd = new MySqlCommand(insertCommand, connection);
                cmd.CommandTimeout = 280000;
                cmd.ExecuteNonQuery();
                connection.Close();
                return true;
            }catch(Exception ex)
            {
                exception = ex.ToString();
                return false;
            }
        }

        public static string GetException(this IMySqlCommand data)
        {
            return exception;
        }


        // not use.
        //private static MySqlConnection GetConnection()
        //{
        //    return new MySqlConnection(DataStore._config.ConnectionString);
        //}
    }
}
