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
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using System.Drawing;
using Microsoft.VisualBasic;

namespace Order_Solutions
{
    class Ayplan
    {
        Calendar calendar = CultureInfo.CurrentCulture.Calendar;
        Table tablo = new Table();
        public string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        public int SameProcess(string whole, String what)
        {
            int count = 0;
            for (int i = 0; i < whole.Length - 1; i++)
            {
                if (whole.Substring(i, 2) == what)
                {
                    count++;
                }
            }
            return count;
        }
        public Ayplan() { }
        public void plan(string plan, string db, string pc, string dbs, string pcs)
        {
            int SN = 1; // Sipariş No
            int BS = 1; // Boya Statü
            int btRowCnt = 1;
            FileInfo file2 = new FileInfo(db);
            FileInfo file1 = new FileInfo(pc);
            using (ExcelPackage dasboard = new ExcelPackage(file2))
            {
                int btColCnt = dasboard.Workbook.Worksheets[dbs].Dimension.End.Column;
                btRowCnt = dasboard.Workbook.Worksheets[dbs].Dimension.End.Row;
                for (int i = 1; i < btColCnt + 1; i++)
                {
                    if (SN != 1 && BS != 1)
                    {
                        break;
                    }

                    if (dasboard.Workbook.Worksheets[dbs].Cells[1, i].Value == null)
                    {
                        break;
                    }
                    else if (dasboard.Workbook.Worksheets[dbs].Cells[1, i].Value.ToString() == "Sipariş No") { SN = i; }
                    else if (dasboard.Workbook.Worksheets[dbs].Cells[1, i].Value.ToString() == "Statü / Termin") { BS = i; }
                }
            }
            FileInfo output = new FileInfo(plan);
            if (output.Exists)
            {
                output.Delete();
                output = new FileInfo(plan);
            }

            using (ExcelPackage ayplani = new ExcelPackage(output))
            {
                ExcelPackage dashboard = new ExcelPackage(file2);
                ExcelPackage plannedcoils = new ExcelPackage(file1);
                var ws2 = ayplani.Workbook.Worksheets.Add("dashboard");
                var ws1 = ayplani.Workbook.Worksheets.Add("aubt", plannedcoils.Workbook.Worksheets[pcs]);

                dashboard.Workbook.Worksheets[dbs].Cells[1, SN, btRowCnt, SN].Copy(ws2.Cells[1, 1, btRowCnt, 1]);
                dashboard.Workbook.Worksheets[dbs].Cells[1, BS, btRowCnt, BS].Copy(ws2.Cells[1, 2, btRowCnt, 2]);
                ayplani.SaveAs(output);

                // satış grubu "BYL" olmayanlar silinecek
                int ColCnt = ws1.Dimension.End.Column;
                int RowCnt = ws1.Dimension.End.Row;
                int takim = 1;
                for (int i = 1; i <= ColCnt; i++)
                {
                    if (ws1.Cells[1, i].Value.ToString() == "Satış Takımı")
                    {
                        takim = i;
                        break;
                    }
                }
                for (int i = 2; i < RowCnt + 1; i++)
                {
                    if (ws1.Cells[i, takim].Value.ToString() != "BYL")
                    {
                        ws1.DeleteRow(i);
                        i--;
                        RowCnt = ws1.Dimension.End.Row;
                    }
                }
                RowCnt = ws1.Dimension.End.Row;
                //var ws3 = ayplani.Workbook.Worksheets.Add("pasifle", ayplani.Workbook.Worksheets["siparisler"]);
                //var ws4 = ayplani.Workbook.Worksheets.Add("Liste");
                ayplani.Save();

                // aubt'ye eklenecek kolonlar 
                int reelmiktar = 1;
                int sure = 1;
                int AV = 1;
                int TP = 1;
                int SH = 1;
                int LT = 1;
                int GD = 1;
                int BY = 1;
                int LD = 1;
                int BK = 1;
                int LP = 1;
                int US = 1; // Üretim Süresi
                int MGS = 1; //Metal Geliş Süre
                int BG = 1; // Boya Geliş
                int RG = 1; // Reel Geliş              
                int GA = 1; //Genel Açıklama
                int BH = 1; // Boyama Haftası
                int KR = 1; // Kalan Rota
                int PM = 1; // Planlanan Miktar
                int UM = 1; // Üretimdeki Miktar
                int SK = 1; // Şipariş Kalınlık
                int GE = 1; // Giriş Eni
                int SP = 1; //Sipariş-Kalem No

                for (int i = 1; i < ColCnt + 1; i++)
                {
                    if (ws1.Cells[1, i].Value.ToString() == "Planlanan Miktar") { PM = i; }
                    if (ws1.Cells[1, i].Value.ToString() == "Sipariş Kalınlık") { SK = i; }
                    if (ws1.Cells[1, i].Value.ToString() == "Giriş En") { GE = i; }
                    if (ws1.Cells[1, i].Value.ToString() == "Üretimdeki Miktar") { UM = i; }
                    if (ws1.Cells[1, i].Value.ToString() == "Kalan Rota") { KR = i; }
                    if (ws1.Cells[1, i].Value.ToString() == "Sipariş-Kalem No") { SP = i; }

                    if (ws1.Cells[1, i].Value.ToString() == "Kullanım Alanı")
                    {
                        ws1.InsertColumn(i + 1, 2, i);
                        ws1.Cells[1, i + 1].Value = "Reel Miktar";
                        ws1.Cells[1, i + 2].Value = "Süre";
                        ws1.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                        ws1.Cells[1, i + 2].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                        reelmiktar = i + 1;
                        sure = i + 2;
                        ColCnt = ws1.Dimension.End.Column;
                    }
                    if (ws1.Cells[1, i].Value.ToString() == "Avalible Time")
                    {
                        ws1.InsertColumn(i, 1, i + 1);
                        ws1.Cells[1, i].Value = "Available";
                        ws1.Cells[1, i].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                        AV = i;
                        ColCnt = ws1.Dimension.End.Column;
                        i++;

                    }
                    if (ws1.Cells[1, i].Value.ToString() == "Kalan Rota")
                    {
                        ws1.InsertColumn(i + 1, 14, i);
                        ws1.Cells[1, i + 1].Value = "TP";
                        ws1.Cells[1, i + 2].Value = "SH";
                        ws1.Cells[1, i + 3].Value = "LT";
                        ws1.Cells[1, i + 4].Value = "GD";
                        ws1.Cells[1, i + 5].Value = "BY";
                        ws1.Cells[1, i + 6].Value = "LD";
                        ws1.Cells[1, i + 7].Value = "BK";
                        ws1.Cells[1, i + 8].Value = "LP";
                        ws1.Cells[1, i + 9].Value = "Üretim Süresi";
                        ws1.Cells[1, i + 10].Value = "Metal Geliş Süre";
                        ws1.Cells[1, i + 11].Value = "Boya Geliş";
                        ws1.Cells[1, i + 12].Value = "Reel Geliş";
                        ws1.Cells[1, i + 13].Value = "Genel Açıklama";
                        ws1.Cells[1, i + 14].Value = "Boyama Hafta";
                        TP = i + 1;
                        SH = i + 2;
                        LT = i + 3;
                        GD = i + 4;
                        BY = i + 5;
                        LD = i + 6;
                        BK = i + 7;
                        LP = i + 8;
                        US = i + 9;
                        MGS = i + 10;
                        BG = i + 11;
                        RG = i + 12;
                        GA = i + 13;
                        BH = i + 14;
                        ColCnt = ws1.Dimension.End.Column;
                        ayplani.Save();
                        using (ExcelRange r = ws1.Cells[1, i + 1, 1, i + 14])
                        {
                            r.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                        }
                    }
                }
                // Reel Miktar Formula

                string rmexcel = GetExcelColumnName(reelmiktar); // Colon numarasını ver Adresini al
                double orthız = double.Parse(Interaction.InputBox("Ortalama Hız", "Ortalama Hızı Giriniz", "75", 0, 0));
                double verim = double.Parse(Interaction.InputBox("Verim", "Ortalama Uptime Giriniz", "0,70", 0, 0));
                DateTime dt = new DateTime();
                for (int i = 2; i < RowCnt + 1; i++)
                {
                    if (double.Parse(ws1.Cells[i, PM].Value.ToString()) == 0 || double.Parse(ws1.Cells[i, UM].Value.ToString()) == 0)
                    {
                        ws1.Cells[i, reelmiktar].Value = Math.Max(double.Parse(ws1.Cells[i, PM].Value.ToString()), double.Parse(ws1.Cells[i, UM].Value.ToString()));
                    }
                    else { ws1.Cells[i, reelmiktar].Value = Math.Min(double.Parse(ws1.Cells[i, PM].Value.ToString()), double.Parse(ws1.Cells[i, UM].Value.ToString())); }
                    ws1.Cells[i, reelmiktar].StyleID = ws1.Cells[i, PM].StyleID;
                    double metraj = double.Parse(ws1.Cells[i, reelmiktar].Value.ToString()) / 2.71 / double.Parse(ws1.Cells[i, SK].Value.ToString()) / (double.Parse(ws1.Cells[i, GE].Value.ToString()) / 1000);
                    ws1.Cells[i, sure].Value = metraj / orthız / verim;
                    ws1.Cells[i, TP].Value = SameProcess(ws1.Cells[i, KR].Value.ToString(), "TP");
                    ws1.Cells[i, SH].Value = SameProcess(ws1.Cells[i, KR].Value.ToString(), "SH");
                    ws1.Cells[i, LT].Value = SameProcess(ws1.Cells[i, KR].Value.ToString(), "LT");
                    ws1.Cells[i, GD].Value = SameProcess(ws1.Cells[i, KR].Value.ToString(), "GD");
                    ws1.Cells[i, BY].Value = SameProcess(ws1.Cells[i, KR].Value.ToString(), "BY");
                    ws1.Cells[i, LD].Value = SameProcess(ws1.Cells[i, KR].Value.ToString(), "LD");
                    ws1.Cells[i, BK].Value = SameProcess(ws1.Cells[i, KR].Value.ToString(), "BK");
                    ws1.Cells[i, LP].Value = SameProcess(ws1.Cells[i, KR].Value.ToString(), "LP");

                    if (ws1.Cells[i, AV + 1].Value.ToString() == "")
                    {
                        ws1.Cells[i, AV].Value = "Planlanacak";
                    }
                    else
                    {
                        ws1.Cells[i, AV].Value = (ws1.Cells[i, AV + 1].Value.ToString().Substring(0, 10));
                        dt = new DateTime(int.Parse(ws1.Cells[i, AV].Value.ToString().Substring(6, 4)), int.Parse(ws1.Cells[i, AV].Value.ToString().Substring(3, 2)), int.Parse(ws1.Cells[i, AV].Value.ToString().Substring(0, 2)));
                        ws1.Cells[i, US].Value =
                        int.Parse(ws1.Cells[i, TP].Value.ToString()) * 2 +
                        int.Parse(ws1.Cells[i, SH].Value.ToString()) * 1 +
                        int.Parse(ws1.Cells[i, LT].Value.ToString()) * 6 +
                        int.Parse(ws1.Cells[i, GD].Value.ToString()) * 1 +
                        int.Parse(ws1.Cells[i, BY].Value.ToString()) * 1;
                        ws1.Cells[i, MGS].Value = (dt.AddDays(int.Parse(ws1.Cells[i, US].Value.ToString())));
                        ws1.Cells[i, MGS].Style.Numberformat.Format = "dd/MM/yyyy";
                    }
                }
                ayplani.Save();
                int RowCnt2 = ws2.Dimension.End.Row;
                int ColCnt2 = ws2.Dimension.End.Column;

                for (int i = 2; i < RowCnt2 + 1; i++)
                {
                    if (ws2.Cells[i, 1].Value.ToString() == "")
                    {
                        break;
                    }
                    else if (ws2.Cells[i, 2].Value.ToString() == "")
                    {
                        ws2.Cells[i, 2].Value = "Termini YOK";
                    }
                    else if (ws2.Cells[i, 2].Value.ToString() != "OK" && ws2.Cells[i, 2].Value.ToString() != "ok")
                    {
                        ws2.Cells[i, 2].Style.Numberformat.Format = "dd/MM/yyyy";
                    }
                }
                ayplani.Save();

                // Sip-Poz'a Göre Boya Terminlerini Doldur
                DateTime dtbg = DateTime.Today;
                DateTime dtmgs = DateTime.Today;
                DateTime dtrg;
                for (int i = 2; i < RowCnt + 1; i++)
                {
                    for (int j = 2; j < RowCnt2 + 1; j++)
                    {

                        if (ws2.Cells[j, 1].Value.ToString() == "")
                        {
                            break;
                        }
                        else if (ws2.Cells[j, 1].Value.ToString() == ws1.Cells[i, SP].Value.ToString())
                        {
                            ws1.Cells[i, BG].Value = ws2.Cells[j, 2].Value;
                            ws1.Cells[i, BG].Style.Numberformat.Format = "dd/MM/yyyy";
                            string orj = ws1.Cells[i, BG].Value.ToString();
                            if (ws1.Cells[i, BG].Value.ToString() == "Termini YOK")
                            {
                                ws1.Cells[i, BG].Value = null;
                            }
                            else if (ws1.Cells[i, BG].Value.ToString().Any(x => char.IsLetter(x)))
                            {
                                ws1.Cells[i, BG].Value = DateTime.Today;
                                ws1.Cells[i, BG].Style.Numberformat.Format = "dd/MM/yyyy";

                            }
                            if (ws1.Cells[i, BG].Value != null)
                            {
                                dtbg = new DateTime(int.Parse(ws1.Cells[i, BG].Value.ToString().Substring(6, 4)), int.Parse(ws1.Cells[i, BG].Value.ToString().Substring(3, 2)), int.Parse(ws1.Cells[i, BG].Value.ToString().Substring(0, 2)));
                            }

                            if (ws1.Cells[i, MGS].Value is null) { }
                            else if (ws1.Cells[i, BG].Value is null) { ws1.Cells[i, BG].Value = orj; }
                            else
                            {
                                dtmgs = new DateTime(int.Parse(ws1.Cells[i, MGS].Value.ToString().Substring(6, 4)), int.Parse(ws1.Cells[i, MGS].Value.ToString().Substring(3, 2)), int.Parse(ws1.Cells[i, MGS].Value.ToString().Substring(0, 2)));
                                int compare = DateTime.Compare(dtbg, dtmgs);
                                if (compare < 0)
                                {
                                    ws1.Cells[i, RG].Value = dtmgs;
                                    ws1.Cells[i, RG].Style.Numberformat.Format = "dd/MM/yyyy";
                                    ws1.Cells[i, BH].Value = (calendar.GetWeekOfYear(dtmgs, CalendarWeekRule.FirstDay, DayOfWeek.Monday) + ".HAFTA");
                                }
                                else
                                {
                                    ws1.Cells[i, RG].Value = dtbg;
                                    ws1.Cells[i, RG].Style.Numberformat.Format = "dd/MM/yyyy";
                                    ws1.Cells[i, BH].Value = (calendar.GetWeekOfYear(dtbg, CalendarWeekRule.FirstDay, DayOfWeek.Monday) + ".HAFTA");
                                }
                                ws1.Cells[i, BG].Value = orj;
                                ws1.Cells[i, BG].Style.Numberformat.Format = "dd/MM/yyyy";

                            }
                            // Genel Açıklama Stünu
                            if (ws1.Cells[i, RG].Value is null) { }
                            else
                            {
                                dtrg = new DateTime(int.Parse(ws1.Cells[i, RG].Value.ToString().Substring(6, 4)), int.Parse(ws1.Cells[i, RG].Value.ToString().Substring(3, 2)), int.Parse(ws1.Cells[i, RG].Value.ToString().Substring(0, 2)));
                                if (calendar.GetMonth(dtrg) == calendar.GetMonth(DateTime.Today) && ws1.Cells[i, BY].Value.ToString() == "1")
                                {
                                    ws1.Cells[i, GA].Value = "ÜRETİLİR";
                                }
                                else if (calendar.GetMonth(dtrg) == calendar.GetMonth(DateTime.Today))
                                {
                                    ws1.Cells[i, GA].Value = "ÜRETİLİR - BOYA SONRASINDA";
                                }
                                else if (calendar.GetMonth(dtrg) > calendar.GetMonth(DateTime.Today))
                                {
                                    if (calendar.GetMonth(dtbg) > calendar.GetMonth(DateTime.Today))
                                    {
                                        ws1.Cells[i, GA].Value = "ÜRETİLMEZ - BOYA GEÇ";
                                    }
                                    else
                                    {
                                        ws1.Cells[i, GA].Value = "ÜRETİLMEZ - METAL GEÇ";
                                    }
                                }
                            }
                            break;
                        }
                    }
                }
                // Son Kontrol Metal Geç Olanlar için
                for (int i = 2; i < RowCnt + 1; i++)
                {
                    if (ws1.Cells[i, MGS].Value is null) { }
                    else
                    {
                        dtmgs = new DateTime(int.Parse(ws1.Cells[i, MGS].Value.ToString().Substring(6, 4)), int.Parse(ws1.Cells[i, MGS].Value.ToString().Substring(3, 2)), int.Parse(ws1.Cells[i, MGS].Value.ToString().Substring(0, 2)));
                        if (calendar.GetMonth(dtmgs) > calendar.GetMonth(DateTime.Today))
                        {
                            ws1.Cells[i, GA].Value = "ÜRETİLMEZ - METAL GEÇ";
                        }
                    }
                }
                // Son Kontrol Boya Sonrasındakiler için
                for (int i = 2; i < RowCnt + 1; i++)
                {
                    if (ws1.Cells[i, BG].Value is null && ws1.Cells[i, BY].Value.ToString() != "1")
                    {
                        ws1.Cells[i, GA].Value = "ÜRETİLİR - BOYA SONRASINDA";
                    }
                }

                // Son Kontrol Termini Yok'lar  için
                for (int i = 2; i < RowCnt + 1; i++)
                {
                    if (ws1.Cells[i, BG].Value is null || ws1.Cells[i, BG].Value.ToString() == "Termini YOK")
                    {
                        if (ws1.Cells[i, GA].Value is null)
                        {
                            ws1.Cells[i, GA].Value = "Boya Termini YOK";
                        }
                    }
                }
                int KP = int.Parse(Interaction.InputBox("KP Miktar", "KP Miktarını giriniz", "75", 0, 0));
                ws1.InsertRow(2, 1, 3);
                for (int i = 1; i < ColCnt+1; i++)
                {
                    if (i == reelmiktar)
                        ws1.Cells[2, reelmiktar].Value = KP;
                    else
                        ws1.Cells[2, i].Value = "KP Üretim";
                }
                ayplani.Save();
            }
        }
    }
}

/*
int order = 1;
int remain = 1;
int uretim = 1;
int colnum = 1;
for (int i = 1; i <= ColCnt; i++)
{
if (ws1.Cells[1, i].Value.ToString() == "Sipariş Miktarı")
{
order = i;
ws4.InsertColumn(i, 2, i - 1);
ws4.Cells[1, i].Value = "Sip-Poz";
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
ayplani.Save();
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
for (int i = 2; i < listrowcount + 1; i++)
{
for (int j = 2; j < aubtrowcount + 1; j++)
{
if (ws4.Cells[i, 4].Value.ToString() == ws2.Cells[j, 2].Value.ToString())
{
ws4.DeleteRow(i, 1, true);
i--;
listrowcount = ws4.Dimension.End.Row;
break;
}
}
}*/
