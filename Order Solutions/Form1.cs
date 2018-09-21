using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
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

namespace Order_Solutions
{
    public partial class Form1 : Form
    {
        /*  all orders = ao;
            all paints = ap
            aübt = pc (planned coils)
            Bobin Takip = ac (all coils)
            Dashboard = DBConcurrencyException */

        string sheetao = "";
        string pathinao = "";
        string sheetap = "";
        string pathinap = "";
        string sheetpc = "";
        string pathinpc = "";
        string pathindb = "";
        string sheetdb = "";
        string sheetac = "";
        string pathinac = "";
        string pathsout = "";
        Combination allsheets = new Combination();
        Passive nestyorders = new Passive();
        Ayplan planmaker = new Ayplan();

        public Form1()
        {
            InitializeComponent();
        }



        //allorders
        private void Callorders_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog ChoosingExcelFile = new OpenFileDialog();
                ChoosingExcelFile.Filter = "Excel Files | *.xlsx; *.xls; *.xlsm; *.xml";
                if (ChoosingExcelFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    this.Pallorders.Text = ChoosingExcelFile.FileName;
                    pathinao = Pallorders.Text;
                }
                string connection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Pallorders.Text + ";Extended Properties = \"Excel 12.0; HDR = YES;\"; ";
                OleDbConnection con = new OleDbConnection(connection);
                con.Open();
                CSallorders.DataSource = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                CSallorders.DisplayMember = "TABLE_NAME";
                CSallorders.ValueMember = "TABLE_NAME";
                con.Close();
                sheetao = CSallorders.SelectedValue.ToString();
                sheetao = sheetao.Substring(0, sheetao.Length - 1);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //allpaints
        private void Callpaints_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog ChoosingExcelFile = new OpenFileDialog();
                ChoosingExcelFile.Filter = "Excel Files | *.xlsx; *.xls; *.xlsm; *.xml";
                if (ChoosingExcelFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    this.Pallpaints.Text = ChoosingExcelFile.FileName;
                    pathinap = Pallpaints.Text;
                }
                string connection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Pallpaints.Text + ";Extended Properties = \"Excel 12.0; HDR = YES;\"; ";
                OleDbConnection con = new OleDbConnection(connection);
                con.Open();
                CSallpaints.DataSource = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                CSallpaints.DisplayMember = "TABLE_NAME";
                CSallpaints.ValueMember = "TABLE_NAME";
                con.Close();
                sheetap = CSallpaints.SelectedValue.ToString();
                sheetap = sheetap.Substring(0, sheetap.Length - 1);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //aubt
        private void Caubt_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog ChoosingExcelFile = new OpenFileDialog();
                ChoosingExcelFile.Filter = "Excel Files | *.xlsx; *.xls; *.xlsm; *.xml";
                if (ChoosingExcelFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    this.Paubt.Text = ChoosingExcelFile.FileName;
                    pathinpc = Paubt.Text;
                }
                string connection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Paubt.Text + ";Extended Properties = \"Excel 12.0; HDR = YES;\"; ";
                OleDbConnection con = new OleDbConnection(connection);
                con.Open();
                CSaubt.DataSource = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                CSaubt.DisplayMember = "TABLE_NAME";
                CSaubt.ValueMember = "TABLE_NAME";
                con.Close();
                sheetpc = CSaubt.SelectedValue.ToString();
                sheetpc = sheetpc.Substring(0, sheetpc.Length - 1);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //bobintakip
        private void Cbobintakip_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog ChoosingExcelFile = new OpenFileDialog();
                ChoosingExcelFile.Filter = "Excel Files | *.xlsx; *.xls; *.xlsm; *.xml";
                if (ChoosingExcelFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    this.Pbobintakip.Text = ChoosingExcelFile.FileName;
                    pathinac = Pbobintakip.Text;
                }
                string connection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Pbobintakip.Text + ";Extended Properties = \"Excel 12.0; HDR = YES;\"; ";
                OleDbConnection con = new OleDbConnection(connection);
                con.Open();
                CSbobintakip.DataSource = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                CSbobintakip.DisplayMember = "TABLE_NAME";
                CSbobintakip.ValueMember = "TABLE_NAME";
                con.Close();
                sheetac = CSbobintakip.SelectedValue.ToString();
                sheetac = sheetac.Substring(0, sheetac.Length - 1);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //dashboard
        private void Cdashboard_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog ChoosingExcelFile = new OpenFileDialog();
                ChoosingExcelFile.Filter = "Excel Files | *.xlsx; *.xls; *.xlsm; *.xml";
                if (ChoosingExcelFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    this.Pdashboard.Text = ChoosingExcelFile.FileName;
                    pathindb = Pdashboard.Text;
                }
                string connection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Pdashboard.Text + ";Extended Properties = \"Excel 12.0; HDR = YES;\"; ";
                OleDbConnection con = new OleDbConnection(connection);
                con.Open();
                CSdashboard.DataSource = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                CSdashboard.DisplayMember = (String)con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null).Rows[0]["TABLE_NAME"];
                CSdashboard.ValueMember = "TABLE_NAME";
                con.Close();
                sheetdb = CSdashboard.SelectedValue.ToString();
                sheetdb = sheetdb.Substring(0, sheetdb.Length - 1);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Passive_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(Pallorders.Text) || (String.IsNullOrEmpty(Paubt.Text)))
            {
                MessageBox.Show("Lütfen Tüm Siparişler ve AÜBT Dosyalarını Seçiniz");
                return;
            }
            string pathsout = pathinao.Substring(0, pathinao.LastIndexOf("."));
            string ext = pathinao.Substring(pathinao.LastIndexOf("."));
            pathsout = pathinpc.Substring(0, pathinpc.LastIndexOf("."));
            string ext2 = pathinpc.Substring(pathinpc.LastIndexOf("."));
            if (ext == ".xls" || ext2 == ".xls")
            {
                MessageBox.Show("Lütfen Dosyaları Excel Yeni Versiyonu ile Kaydediniz");
            }
            var dialog = new FolderBrowserDialog();
            dialog.ShowDialog();
            pathsout = dialog.SelectedPath;
            pathsout = pathsout + "\\Pasifle" + ".xlsx";
            nestyorders.forpassive(pathsout, pathinao, pathinpc, sheetao, sheetpc);
        }



        private void Combination_Click(object sender, EventArgs e)
        {
            var dialog = new FolderBrowserDialog();
            dialog.ShowDialog();
            pathsout = dialog.SelectedPath;
            pathsout = pathsout + "\\combo" + ".xlsx";
            allsheets.oneexcel(pathsout, pathinao, pathinap, pathinpc, pathinac);
        }

        private void ayplan_Click(object sender, EventArgs e)
        {


            if (String.IsNullOrEmpty(Pdashboard.Text) || (String.IsNullOrEmpty(Paubt.Text)))
            {
                MessageBox.Show("Lütfen DAshBoard ve AÜBT Dosyalarını Seçiniz");
                return;
            }
            // BUraya "'" ve "$" ları silecek şekilde formül yaz.
            sheetdb = CSdashboard.SelectedValue.ToString();
            sheetdb = Utilities.Trimsheetnames(sheetdb);
            sheetdb = sheetdb.Substring(1, sheetdb.Length - 3);
            sheetpc = CSaubt.SelectedValue.ToString();
            sheetpc = sheetpc.Substring(0, sheetpc.Length - 1);
            string pathsout = pathindb.Substring(0, pathindb.LastIndexOf("."));
            string ext = pathindb.Substring(pathindb.LastIndexOf("."));
            pathsout = pathinpc.Substring(0, pathinpc.LastIndexOf("."));
            string ext2 = pathinpc.Substring(pathinpc.LastIndexOf("."));
            if (ext == ".xls" || ext2 == ".xls")
            {
                MessageBox.Show("Lütfen Dosyaları Excel Yeni Versiyonu ile Kaydediniz");
            }
            var dialog = new FolderBrowserDialog();
            dialog.ShowDialog();
            pathsout = dialog.SelectedPath;
            pathsout = pathsout + "\\AYPLANI" + ".xlsx";
            planmaker.plan(pathsout, pathindb, pathinpc, sheetdb, sheetpc);
        }
    }
}
