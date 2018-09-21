using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using OfficeOpenXml;
using System.IO;
using System.Globalization;
using System.Xml;
using System.Text.RegularExpressions;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Style;
using OfficeOpenXml.VBA;

namespace Order_Solutions
{
    // Combine All Sheets
    class Combination
    {
        public Combination() { }

        public void oneexcel(string combo, string ao, string ap, string pc, string ac)
        {
            FileInfo file1 = new FileInfo(ao);
            FileInfo file2 = new FileInfo(ap);
            FileInfo file3 = new FileInfo(pc);
            FileInfo file4 = new FileInfo(ac);
            FileInfo output = new FileInfo(combo);
            if (output.Exists)
            {
                output.Delete();
                output = new FileInfo(combo);
            }

            using (ExcelPackage Combo = new ExcelPackage(output))
            {
                ExcelPackage allorders = new ExcelPackage(file1);
                ExcelPackage allpaints = new ExcelPackage(file2);
                ExcelPackage plannedcoils = new ExcelPackage(file3);
                ExcelPackage allcoils = new ExcelPackage(file4);
                var ws1 = Combo.Workbook.Worksheets.Add("siparisler", allorders.Workbook.Worksheets["Sheet1"]);
                var ws2 = Combo.Workbook.Worksheets.Add("boyalar", allpaints.Workbook.Worksheets["Sheet1"]);
                var ws3 = Combo.Workbook.Worksheets.Add("aubt", plannedcoils.Workbook.Worksheets["Sheet1"]);
                var ws4 = Combo.Workbook.Worksheets.Add("bobintakip", allcoils.Workbook.Worksheets["Sheet1"]);
                Combo.SaveAs(output);

                // AO'dan toplam satırları silme
                int ColCnt = ws1.Dimension.End.Column;
                int RowCnt = ws1.Dimension.End.Row;
                for (int i = 1; i < RowCnt + 1; i++)
                {
                    if (ws1.Cells[i, 1].Value == null) { break; }

                    else if (ws1.Cells[i, 1].Value.ToString() == "Toplam" || ws1.Cells[i, 2].Value.ToString() == "Toplam")
                    {
                        ws1.DeleteRow(i);
                        i--;
                    }
                }
                RowCnt = ws1.Dimension.End.Row;
                Combo.Save();
                

                //Boyaları özet tablo yapma
                
            }
        }
        
    }
}
