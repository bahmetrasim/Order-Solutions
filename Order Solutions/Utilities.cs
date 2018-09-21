using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Order_Solutions
{
    class Utilities
    {
        // Split Methodunu yaz 
        // tek tırnak char // çift tırnak string 
        public static string Trimsheetnames (string Sheetname)
        {
            if (Sheetname.Split(' ').Length == 1)
            {
                Sheetname = Sheetname.Substring(0, Sheetname.Length - 1);
            }
            else
            {
                Sheetname = Sheetname.Substring(1, Sheetname.Length - 3);
            }

            return Sheetname;
        }

    }
}
