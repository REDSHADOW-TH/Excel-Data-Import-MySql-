//Copyright 2020 REDSHADOWTH(Phumiphat Joreonyonwhatthana) under MIT Licence.
//MIT Licence https://opensource.org/licenses/MIT

using System;
using System.Collections.Generic;
using System.Text;

namespace Excel_Data_Import__MySql_.Data
{
    public struct IExcelConfig
    {
        public string Path;
        public string Sheet;
        public int StartRow;
    }
}
