//Copyright 2020 REDSHADOWTH(Phumiphat Joreonyonwhatthana) under MIT Licence.
//MIT Licence https://opensource.org/licenses/MIT

using System;
using System.IO;
using ironLua = Neo.IronLua;
using Excel_Data_Import__MySql_.Data;

namespace Excel_Data_Import__MySql_
{
    public static class Config
    {
        public static IConfig GetConfig()
        {
            IConfig result = new IConfig() { ConnectionString = "", ScriptPath = "" };
            string configPath = @$"{Directory.GetCurrentDirectory()}\config.lua";

            if (!File.Exists(configPath))
            {
                Console.WriteLine("Can not found config.lua please check and try again !");
            }
            else
            {
                string luaScript = "";
                using (StreamReader file = new StreamReader(configPath))
                {
                    string readResult = file.ReadToEnd();
                    if (String.IsNullOrEmpty(readResult))
                    {
                        Console.WriteLine("config is valid !");
                    } else
                    {
                        luaScript = readResult;
                    }
                }
                ironLua.Lua _lua = new ironLua.Lua();
                dynamic env = _lua.CreateEnvironment();

                env.dochunk(luaScript);

                if (env.config != null)
                {
                    result = new IConfig()
                    {
                        ConnectionString = env.config.connectionString ?? "",
                        ScriptPath = env.config.scriptPath ?? ""
                    };
                }
            }
            return result;
        }
    }
}
