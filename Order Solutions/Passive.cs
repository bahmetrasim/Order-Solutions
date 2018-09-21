using System;
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
    class Passive
    {
        public Passive() { }
        public void passivation(string combo, string ao, string pc)
        {

        }
        public void forpassive(string pasif, string ao, string pc, string aos, string pcs)
        {

            FileInfo file1 = new FileInfo(ao);
            FileInfo file2 = new FileInfo(pc);
            FileInfo output = new FileInfo(pasif);
            if (output.Exists)
            {
                output.Delete();
                output = new FileInfo(pasif);
            }

            using (ExcelPackage Pasiflenecek = new ExcelPackage(output))
            {
                ExcelPackage allorders = new ExcelPackage(file1);
                ExcelPackage plannedcoils = new ExcelPackage(file2);
                var ws1 = Pasiflenecek.Workbook.Worksheets.Add("siparisler", allorders.Workbook.Worksheets[aos]);
                var ws2 = Pasiflenecek.Workbook.Worksheets.Add("aubt", plannedcoils.Workbook.Worksheets[pcs]);
                Pasiflenecek.SaveAs(output);

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
                var ws3 = Pasiflenecek.Workbook.Worksheets.Add("pasifle", Pasiflenecek.Workbook.Worksheets["siparisler"]);
                var ws4 = Pasiflenecek.Workbook.Worksheets.Add("Liste");
                Pasiflenecek.Save();

                // AO'da %20 Tolerans Kontrol 
                int order = 1;
                int remain = 1;
                int uretim = 1;
                int colnum = 1;
                for (int i = 1; i <= ColCnt; i++)
                {
                    if (ws1.Cells[1, i].Value.ToString() == "Sipariş Miktarı")
                    {
                        order = i;
                    }
                    if (ws1.Cells[1, i].Value.ToString() == "Kalan Miktar")
                    {
                        remain = i;
                    }
                    if (ws1.Cells[1, i].Value.ToString() == "Üretilen Miktar")
                    {
                        uretim = i;
                    }
                    if (order != 1 && remain != 1 && uretim != 1)
                        break;
                }
                ws3.Cells[colnum, 1, 1, remain].Copy(ws4.Cells[1, 1, 1, remain]);

                for (int i = 2; i < RowCnt + 1; i++)
                {
                    if ((double.Parse(ws3.Cells[i, remain].Value.ToString())) / double.Parse(ws3.Cells[i, order].Value.ToString()) <= 0.20)
                    {
                        colnum++;
                        ws3.Cells[i, 1, i, remain].Copy(ws4.Cells[colnum, 1, i, remain]);
                        ws3.DeleteRow(i, 1, true);
                        i--;
                        RowCnt = ws3.Dimension.End.Row;
                    }
                }
                Pasiflenecek.Save();
                // Miktar 2 Ton altı olanları Kontrol et
                for (int i = 2; i < RowCnt + 1; i++)
                {
                    if ((double.Parse(ws3.Cells[i, remain].Value.ToString())) < 2)
                    {
                        colnum++;
                        ws3.Cells[i, 1, i, remain].Copy(ws4.Cells[colnum, 1, i, remain]);
                        ws3.DeleteRow(i, 1, true);
                        i--;
                        RowCnt = ws3.Dimension.End.Row;
                    }
                }
                int listrowcount = ws4.Dimension.End.Row;
                for (int i = 2; i < listrowcount + 1; i++)
                {
                    if ((double.Parse(ws4.Cells[i, uretim].Value.ToString())) == 0)
                    {
                        ws4.DeleteRow(i, 1, true);
                        i--;
                        listrowcount = ws4.Dimension.End.Row;
                    }

                }
                // Listeye ek kolon ekleyerek Müs.sip.poz ekle
                listrowcount = ws4.Dimension.End.Row;
                ws4.InsertColumn(4, 1, 3);
                ws4.Cells[1, 4].Value = "Sip-Poz";
                for (int i = 2; i < listrowcount + 1; i++)
                {
                    ws4.Cells[i, 4].Value = ws4.Cells[i, 2].Value.ToString() + "-" + int.Parse(ws4.Cells[i, 3].Value.ToString()).ToString();
                }

                // Aubt ile Karşılaştır - Planlanmış Bobinler var ise Sarıya işaretle

                int aubtrowcount = ws2.Dimension.End.Row;
                for (int i = 2; i < listrowcount +1; i++)
                {
                    for (int j = 2; j <aubtrowcount + 1; j++)
                    {
                        if (ws4.Cells[i,4].Value.ToString() == ws2.Cells[j,2].Value.ToString())
                        {
                            ws4.DeleteRow(i, 1, true);
                            i--;
                            listrowcount = ws4.Dimension.End.Row;
                            break;
                        }
                    }
                }
                Pasiflenecek.Save();
            }
        }
    }
}

