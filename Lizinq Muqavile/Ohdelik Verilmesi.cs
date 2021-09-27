using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Xml;
using System.IO;
using System.Net;
using System.Globalization;
using Nsoft;

namespace Lizinq_Muqavile
{
    public partial class Ohdelik_Verilmesi : Form
    {

        Excel.Application oXL;
        Excel._Workbook oWB;
        Excel._Worksheet oSheet;

        public Ohdelik_Verilmesi()
        {
            InitializeComponent();
        }

        public void Ohdelik()
        {
            try { File.Copy("Ohdelik.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Ohdeliklerin verilmesi - " + txtOhdelikQebulEden.Text + ".doc", true); }
            catch { MessageBox.Show("'\\Ohdelik.doc' tapılmadı."); }

            object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\Ohdeliklerin verilmesi - " + txtOhdelikQebulEden.Text + ".doc";

            Word.Application word = new Word.Application();
            Word.Document doc = null;
            object missing = System.Type.Missing;
            object readOnly = false;
            object isVisible = false;
            word.Visible = true;

            doc = word.Documents.Open(ref FileName);
            doc.Activate();

            DateTime dt2 = txtmuqTARIX.Value.Date;
            DateTime dt = txthazirkiTARIX.Value.Date;

            string a = MyChange.TarixSozle(dt2);
            string b = MyChange.TarixSozle(dt);
            string c = MyChange.ReqemToMetn(Convert.ToInt32(txtmuddet.Text));

            MyChange.FindAndReplace(word, "000", txtOhdelikNo.Text);
            MyChange.FindAndReplace(word, "000", txtOhdelikNo.Text);
            MyChange.FindAndReplace(word, "000", txtOhdelikNo.Text);
            MyChange.FindAndReplace(word, "000", txtOhdelikNo.Text);
            MyChange.FindAndReplace(word, "000", txtOhdelikNo.Text);
            MyChange.FindAndReplace(word, "000", txtOhdelikNo.Text);
            MyChange.FindAndReplace(word, "000", txtOhdelikNo.Text);
            MyChange.FindAndReplace(word, "000", txtOhdelikNo.Text);
            MyChange.FindAndReplace(word, "111", dt.Day + " " + b + " " + dt.Year + "- ci il");
            MyChange.FindAndReplace(word, "111", dt.Day + " " + b + " " + dt.Year + "- ci il");
            MyChange.FindAndReplace(word, "111", dt.Day + " " + b + " " + dt.Year + "- ci il");
            MyChange.FindAndReplace(word, "111", dt.Day + " " + b + " " + dt.Year + "- ci il");
            MyChange.FindAndReplace(word, "111", dt.Day + " " + b + " " + dt.Year + "- ci il");
            MyChange.FindAndReplace(word, "111", dt.Day + " " + b + " " + dt.Year + "- ci il");
            MyChange.FindAndReplace(word, "111", dt.Day + " " + b + " " + dt.Year + "- ci il");
            MyChange.FindAndReplace(word, "222", txtOhdelikveren.Text);
            MyChange.FindAndReplace(word, "222", txtOhdelikveren.Text);
            MyChange.FindAndReplace(word, "222", txtOhdelikveren.Text);
            MyChange.FindAndReplace(word, "222", txtOhdelikveren.Text);
            MyChange.FindAndReplace(word, "222", txtOhdelikveren.Text);
            MyChange.FindAndReplace(word, "222", txtOhdelikveren.Text);
            MyChange.FindAndReplace(word, "222", txtOhdelikveren.Text);
            MyChange.FindAndReplace(word, "222", txtOhdelikveren.Text);
            MyChange.FindAndReplace(word, "333", txtOhdelikQebulEden.Text);
            MyChange.FindAndReplace(word, "333", txtOhdelikQebulEden.Text);
            MyChange.FindAndReplace(word, "333", txtOhdelikQebulEden.Text);
            MyChange.FindAndReplace(word, "333", txtOhdelikQebulEden.Text);
            MyChange.FindAndReplace(word, "333", txtOhdelikQebulEden.Text);
            MyChange.FindAndReplace(word, "333", txtOhdelikQebulEden.Text);
            MyChange.FindAndReplace(word, "333", txtOhdelikQebulEden.Text);
            MyChange.FindAndReplace(word, "333", txtOhdelikQebulEden.Text);
            MyChange.FindAndReplace(word, "333", txtOhdelikQebulEden.Text);
            MyChange.FindAndReplace(word, "333", txtOhdelikQebulEden.Text);
            MyChange.FindAndReplace(word, "333", txtOhdelikQebulEden.Text);
            MyChange.FindAndReplace(word, "333", txtOhdelikQebulEden.Text);
            MyChange.FindAndReplace(word, "444", dt2.Day + " " + a + " " + dt2.Year + "- ci il");
            MyChange.FindAndReplace(word, "444", dt2.Day + " " + a + " " + dt2.Year + "- ci il");
            MyChange.FindAndReplace(word, "444", dt2.Day + " " + a + " " + dt2.Year + "- ci il");
            MyChange.FindAndReplace(word, "444", dt2.Day + " " + a + " " + dt2.Year + "- ci il");
            MyChange.FindAndReplace(word, "555", txtMuqNOMRESI.Text);
            MyChange.FindAndReplace(word, "555", txtMuqNOMRESI.Text);
            MyChange.FindAndReplace(word, "555", txtMuqNOMRESI.Text);
            MyChange.FindAndReplace(word, "555", txtMuqNOMRESI.Text);
            MyChange.FindAndReplace(word, "666", txtobyekt.Text);
            MyChange.FindAndReplace(word, "666", txtobyekt.Text);
            MyChange.FindAndReplace(word, "666", txtobyekt.Text);
            MyChange.FindAndReplace(word, "666", txtobyekt.Text);
            MyChange.FindAndReplace(word, "777", txtmuddet.Text + " (" + c + ")");
            MyChange.FindAndReplace(word, "888", txtumumi.Text + " (" + txtherf1.Text + ")");

            doc.Save();
        }

        public void Elave1()
        {
            try { File.Copy("OhdelikElave1.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Əlavə 1 - " + txtOhdelikQebulEden.Text + ".doc", true); }
            catch { MessageBox.Show("'\\OhdelikElave1.doc' tapılmadı."); }

            object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\Əlavə 1 - " + txtOhdelikQebulEden.Text + ".doc";

            Word.Application word = new Word.Application();
            Word.Document doc = null;
            object missing = System.Type.Missing;
            object readOnly = false;
            object isVisible = false;
            word.Visible = true;

            doc = word.Documents.Open(ref FileName);
            doc.Activate();

            DateTime dt2 = txtmuqTARIX.Value.Date;
            DateTime dt = txthazirkiTARIX.Value.Date;

            string a = MyChange.TarixSozle(dt2);
            string b = MyChange.TarixSozle(dt);

            MyChange.FindAndReplace(word, "000", txtOhdelikNo.Text);
            MyChange.FindAndReplace(word, "000", txtOhdelikNo.Text);
            MyChange.FindAndReplace(word, "000", txtOhdelikNo.Text);
            MyChange.FindAndReplace(word, "000", txtOhdelikNo.Text);
            MyChange.FindAndReplace(word, "000", txtOhdelikNo.Text);
            MyChange.FindAndReplace(word, "000", txtOhdelikNo.Text);
            MyChange.FindAndReplace(word, "000", txtOhdelikNo.Text);
            MyChange.FindAndReplace(word, "111", dt.Day + " " + b + " " + dt.Year + "- ci il");
            MyChange.FindAndReplace(word, "111", dt.Day + " " + b + " " + dt.Year + "- ci il");
            MyChange.FindAndReplace(word, "111", dt.Day + " " + b + " " + dt.Year + "- ci il");
            MyChange.FindAndReplace(word, "111", dt.Day + " " + b + " " + dt.Year + "- ci il");
            MyChange.FindAndReplace(word, "111", dt.Day + " " + b + " " + dt.Year + "- ci il");
            MyChange.FindAndReplace(word, "222", txtOhdelikveren.Text);
            MyChange.FindAndReplace(word, "222", txtOhdelikveren.Text);
            MyChange.FindAndReplace(word, "333", txtOhdelikQebulEden.Text);
            MyChange.FindAndReplace(word, "333", txtOhdelikQebulEden.Text);
            MyChange.FindAndReplace(word, "444", dt2.Day + " " + a + " " + dt2.Year + "- ci il");
            MyChange.FindAndReplace(word, "444", dt2.Day + " " + a + " " + dt2.Year + "- ci il");
            MyChange.FindAndReplace(word, "444", dt2.Day + " " + a + " " + dt2.Year + "- ci il");
            MyChange.FindAndReplace(word, "444", dt2.Day + " " + a + " " + dt2.Year + "- ci il");
            MyChange.FindAndReplace(word, "555", txtMuqNOMRESI.Text);
            MyChange.FindAndReplace(word, "555", txtMuqNOMRESI.Text);
            MyChange.FindAndReplace(word, "555", txtMuqNOMRESI.Text);
            MyChange.FindAndReplace(word, "555", txtMuqNOMRESI.Text);
            MyChange.FindAndReplace(word, "666", txtobyekt.Text);
            MyChange.FindAndReplace(word, "5555", txtnomre.Text);
            MyChange.FindAndReplace(word, "6666", txtsehadetname.Text);
            MyChange.FindAndReplace(word, "7777", txtsassi.Text);
            MyChange.FindAndReplace(word, "8888", txtban.Text);
            MyChange.FindAndReplace(word, "9999", txtmuherrik.Text);
            MyChange.FindAndReplace(word, "11111", txtrengi.Text);
            MyChange.FindAndReplace(word, "22222", txtburaxilis.Text);

            doc.Save();
        }

        public void PrintYekunSifaris()
        {

            try { File.Copy("Yekun sifariş.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Yekun sifariş - " + txtOhdelikQebulEden.Text + ".doc", true); }
            catch { MessageBox.Show("'Yekun sifariş.doc' tapılmadı."); }

            object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\Yekun sifariş - " + txtOhdelikQebulEden.Text + ".doc";

            Word.Application word = new Word.Application();
            Word.Document doc = null;
            object missing = System.Type.Missing;
            object readOnly = false;
            object isVisible = false;
            word.Visible = true;

            doc = word.Documents.Open(ref FileName);
            doc.Activate();

            DateTime dt = txthazirkiTARIX.Value.Date;
            string b = MyChange.TarixSozle(dt);

            MyChange.FindAndReplace(word, "000", txtOhdelikNo.Text);
            MyChange.FindAndReplace(word, "111", dt.Day + " " + b + " " + dt.Year + "- cı il");
            MyChange.FindAndReplace(word, "222", txtOhdelikQebulEden.Text);
            MyChange.FindAndReplace(word, "222", txtOhdelikQebulEden.Text + " (" + txtSeriya.Text + " № " + txtSeriyaNomre.Text + ")");
            MyChange.FindAndReplace(word, "222", txtOhdelikQebulEden.Text);
            MyChange.FindAndReplace(word, "222", txtOhdelikQebulEden.Text);
            MyChange.FindAndReplace(word, "333", txtobyekt.Text);
            MyChange.FindAndReplace(word, "444", txtavadanliqdeyer.Text + " (" + txtherf2.Text + ")");
            MyChange.FindAndReplace(word, "555", txtOhdelikveren.Text);
            MyChange.FindAndReplace(word, "666", "0 (sıfır)");
            MyChange.FindAndReplace(word, "777", txtUnvan.Text);
            MyChange.FindAndReplace(word, "888", tel1.Text);
            MyChange.FindAndReplace(word, "888", tel1.Text);
            MyChange.FindAndReplace(word, "888", tel1.Text);
            MyChange.FindAndReplace(word, "999", tel2.Text);
            MyChange.FindAndReplace(word, "999", tel2.Text);
            MyChange.FindAndReplace(word, "999", tel2.Text);

            doc.Save();
        }

        public void PrintQrafik()                //--------------------Print qrafik---------------------------------------
        {
            try { File.Copy("Qrafik.xlsm", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Qrafik - " + txtOhdelikQebulEden.Text + ".xlsm", true); }
            catch { MessageBox.Show("'Yekun sifariş.doc' tapılmadı."); }

            DateTime dt = txthazirkiTARIX.Value.Date;
            DateTime dt2 = txtmuqTARIX.Value.Date;

            int s = 0, s3 = 0, s2s = 0, s3s = 0, s4s = 0, s5s = 0;

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Qrafik - " + txtOhdelikQebulEden.Text + ".xlsm"));
            oSheet = (Excel._Worksheet)oWB.Sheets[2];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

            oSheet.Cells[5, 3] = txtOhdelikNo.Text;
            oSheet.Cells[6, 3] = dt.ToString("dd.MM.yy");

            if (txtmuddet.Text == "12") oSheet.Cells[13, 3] = dt.AddMonths(12).ToString("dd.MM.yy");
            else if (txtmuddet.Text == "24") oSheet.Cells[13, 3] = dt.AddMonths(24).ToString("dd.MM.yy");
            else if (txtmuddet.Text == "36") oSheet.Cells[13, 3] = dt.AddMonths(36).ToString("dd.MM.yy");
            else if (txtmuddet.Text == "48") oSheet.Cells[13, 3] = dt.AddMonths(48).ToString("dd.MM.yy");
            else if (txtmuddet.Text == "60") oSheet.Cells[13, 3] = dt.AddMonths(60).ToString("dd.MM.yy");

            oSheet.Cells[7, 3] = txtmuddet.Text;
            oSheet.Cells[8, 3] = (Convert.ToDouble(txtfaiz.Text) / 100).ToString();
            oSheet.Cells[10, 3] = "0";

            oSheet.Cells[12, 3] = dt.AddMonths(1).ToString("dd.MM.yy");
            //else if (Convert.ToInt32(k1) != 12 && Convert.ToInt32(k1) + 1 < 10) oSheet.Cells[12, 3] = txthazirkiTARIX.Text.Substring(0, 2) + ".0" + (Convert.ToInt32(k1) + 1).ToString() + "." + k2;
            //else if (Convert.ToInt32(k1) != 12 && Convert.ToInt32(k1) + 1 > 9) oSheet.Cells[12, 3] = txthazirkiTARIX.Text.Substring(0, 2) + "." + (Convert.ToInt32(k1) + 1).ToString() + "." + k2;

            oSheet.Cells[15, 3] = (Convert.ToDouble(txtavans.Text) * 100 / Convert.ToDouble(txtumumi.Text)/100).ToString();
            oSheet.Cells[16, 3] = "0";
            oSheet.Cells[21, 3] = txtumumi.Text;
            oSheet.Cells[30, 3] = "0";
            oSheet.Cells[38, 3] = txtsigorta.Text;

            for (s = 0; s < txtOhdelikQebulEden.Text.Length; s++)
            {
                if (txtOhdelikQebulEden.Text.Substring(s, 1) == " ") s3 = s3 + 1;
                if (s3 == 0) s2s = s;
                else if (s3 == 1) s3s = s;
                else if (s3 == 2) s4s = s;
                else if (s3 == 3) s5s = s;
            }
            try
            {
                oSheet.Cells[51, 3] = txtOhdelikQebulEden.Text.Substring(s2s + 2, s3s - s2s - 1);
                oSheet.Cells[52, 3] = txtOhdelikQebulEden.Text.Substring(0, s2s + 1);
                oSheet.Cells[53, 3] = txtOhdelikQebulEden.Text.Substring(s3s + 2, s4s - s3s - 1);
            }
            catch { };

            //**************************monitorinqin tarixinin yazilması*************************************************
            oSheet = (Excel._Worksheet)oWB.Sheets[3];
            oSheet.Activate();
            oSheet.Range["A1"].Select();

            oSheet.Cells[1, 1] = "Ödəmə qrafiki: əsas məbləğ və mükafat daxil edilməklə  " + txtMuqNOMRESI.Text + " saylı " + dt2.Day + "." + dt2.Month + "." + dt2.Year + " tarixli öhdəliklərin verilməsi müqaviləsinə Əlavə 3";
            oXL.Visible = true;
            oXL.DisplayAlerts = false;
            oWB.Save();
        }

        public void QebulIstifade()
        {
            try { File.Copy("OhdelikQebul-istifade.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Qebul istifadə - " + txtOhdelikQebulEden.Text + ".doc", true); }
            catch { MessageBox.Show("'\\OhdelikQebul-istifade.doc' tapılmadı."); }

            object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\Qebul istifadə - " + txtOhdelikQebulEden.Text + ".doc";

            Word.Application word = new Word.Application();
            Word.Document doc = null;
            object missing = System.Type.Missing;
            object readOnly = false;
            object isVisible = false;
            word.Visible = true;

            doc = word.Documents.Open(ref FileName);
            doc.Activate();

            DateTime dt = txthazirkiTARIX.Value.Date;
            DateTime dt2 = txthazirkiTARIX.Value.Date;

            string a = MyChange.TarixSozle(dt2);
            string b = MyChange.TarixSozle(dt);

            MyChange.FindAndReplace(word, "000", txtOhdelikNo.Text);
            MyChange.FindAndReplace(word, "000", txtOhdelikNo.Text);
            MyChange.FindAndReplace(word, "000", txtOhdelikNo.Text);
            MyChange.FindAndReplace(word, "000", txtOhdelikNo.Text);
            MyChange.FindAndReplace(word, "111", dt.Day + " " + b + " " + dt.Year + "- ci il");
            MyChange.FindAndReplace(word, "111", dt.Day + " " + b + " " + dt.Year + "- ci il");
            MyChange.FindAndReplace(word, "111", dt.Day + " " + b + " " + dt.Year + "- ci il");
            MyChange.FindAndReplace(word, "111", dt.Day + " " + b + " " + dt.Year + "- ci il");
            MyChange.FindAndReplace(word, "111", dt.Day + " " + b + " " + dt.Year + "- ci il");
            MyChange.FindAndReplace(word, "111", dt.Day + " " + b + " " + dt.Year + "- ci il");
            MyChange.FindAndReplace(word, "222", txtOhdelikveren.Text);
            MyChange.FindAndReplace(word, "222", txtOhdelikveren.Text);
            MyChange.FindAndReplace(word, "222", txtOhdelikveren.Text);
            MyChange.FindAndReplace(word, "333", txtOhdelikQebulEden.Text);
            MyChange.FindAndReplace(word, "333", txtOhdelikQebulEden.Text);
            MyChange.FindAndReplace(word, "333", txtOhdelikQebulEden.Text);
            MyChange.FindAndReplace(word, "444", dt2.Day + " " + a + " " + dt2.Year + "- ci il");
            MyChange.FindAndReplace(word, "555", txtMuqNOMRESI.Text);
            MyChange.FindAndReplace(word, "666", txtobyekt.Text);
            MyChange.FindAndReplace(word, "888", txtavadanliqdeyer.Text + " (" + txtherf2.Text + ")");
            MyChange.FindAndReplace(word, "888", txtavadanliqdeyer.Text + " (" + txtherf2.Text + ")");
            MyChange.FindAndReplace(word, "888", txtavadanliqdeyer.Text + " (" + txtherf2.Text + ")");
            MyChange.FindAndReplace(word, "5555", txtnomre.Text);
            MyChange.FindAndReplace(word, "6666", txtsehadetname.Text);
            MyChange.FindAndReplace(word, "7777", txtsassi.Text);
            MyChange.FindAndReplace(word, "8888", txtban.Text);
            MyChange.FindAndReplace(word, "9999", txtmuherrik.Text);
            MyChange.FindAndReplace(word, "11111", txtrengi.Text);
            MyChange.FindAndReplace(word, "22222", txtburaxilis.Text);

            doc.Save();
        }

        public void reqemler()      //------reqem yazi ile---------------------------------------------------------------
        {
            try { txtherf1.Text = MyChange.ReqemToMetn(Convert.ToDouble(txtumumi.Text)); } catch { }
        }

        public void reqemler2()      //------reqem yazi ile---------------------------------------------------------------
        {
            try { txtherf2.Text = MyChange.ReqemToMetn(Convert.ToDouble(txtavadanliqdeyer.Text)); } catch { }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Ohdelik();
        }

        private void əlaqəToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Elaqe elaqe = new Elaqe();
            elaqe.ShowDialog();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Elave1();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            QebulIstifade();
        }
        private void txtnomre_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    MyData.selectCommand("baza.accdb", "Select * from etibarnameneqliyyat where c1 Like '%" + txtnomre.Text + "%'");
                    MyData.dtmain = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmain);

                    DateTime dt = txthazirkiTARIX.Value.Date;

                    txtOhdelikveren.Text = MyData.dtmain.Rows[0]["c3"].ToString();
                    txtMuqNOMRESI.Text = MyData.dtmain.Rows[0]["c4"].ToString();
                    txtOhdelikNo.Text = txtMuqNOMRESI.Text.Substring(0,1) + "-01/" + dt.Year.ToString().Substring(2, 2) + "-" + MyData.dtmain.Rows[0]["c4"].ToString().Substring(2, MyData.dtmain.Rows[0]["c4"].ToString().Length - 2);
                    txtobyekt.Text = "1 ədəd " + MyData.dtmain.Rows[0]["c2"].ToString() + " markalı avtomobil";
                    //txtavadanliq.Text = MyData.dtmain.Rows[0][5].ToString();
                    //txtfaiz.Text = MyData.dtmain.Rows[0][8].ToString();
                    //txtmuqTARIX.Text = MyData.dtmain.Rows[0][19].ToString();
                    //txtmuddet.Text = MyData.dtmain.Rows[0][6].ToString();
                    txtnomre.Text = MyData.dtmain.Rows[0]["c1"].ToString();
                    txtsehadetname.Text = MyData.dtmain.Rows[0]["c12"].ToString();
                    txtsassi.Text = MyData.dtmain.Rows[0]["c10"].ToString();
                    txtban.Text = MyData.dtmain.Rows[0]["c8"].ToString();
                    txtmuherrik.Text = MyData.dtmain.Rows[0]["c9"].ToString();
                    txtrengi.Text = MyData.dtmain.Rows[0]["c5"].ToString();
                    txtburaxilis.Text = MyData.dtmain.Rows[0]["c6"].ToString();

                    MyData.selectCommand("baza.accdb", "Select * from muqavilelayihe where [Lizinq alan] Like " + "'%" + txtOhdelikveren.Text + "%'");
                    MyData.dtmain = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmain);

                    txtavadanliqdeyer.Text = MyData.dtmain.Rows[0]["Avadanlığın dəyəri"].ToString();
                    txtfaiz.Text = MyData.dtmain.Rows[0]["% dərəcəsi"].ToString();
                    txtmuqTARIX.Text = MyData.dtmain.Rows[0]["Tarix"].ToString();
                    txtmuddet.Text = MyData.dtmain.Rows[0]["Lizinqin müddəti"].ToString();
                }

            }
            catch { };
        }

        private void txtumumi_TextChanged(object sender, EventArgs e)
        {
            reqemler();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Ohdelik();
            Elave1();
            PrintYekunSifaris();
            QebulIstifade();
            PrintQrafik();
        }

        private void txtOhdelikQebulEden_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    MyData.selectCommand("baza.accdb", "Select * from etibarnamesurucu where a1 Like " + "'%" + txtOhdelikQebulEden.Text + "%'");
                    MyData.dtmain = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmain);

                    txtOhdelikQebulEden.Text = MyData.dtmain.Rows[0]["a1"].ToString();
                    txtSeriya.Text = MyData.dtmain.Rows[0]["a2"].ToString().Substring(0, 3);
                    txtSeriyaNomre.Text = MyData.dtmain.Rows[0]["a2"].ToString().Substring(6, MyData.dtmain.Rows[0]["a2"].ToString().Length - 6);
                    txtUnvan.Text = MyData.dtmain.Rows[0]["a4"].ToString();

                    MyData.selectCommand("baza.accdb", "Select * from Telefon where c1 Like '%" + txtOhdelikQebulEden.Text + "%'");
                    MyData.dtmain = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmain);

                    tel1.Text = MyData.dtmain.Rows[0]["c2"].ToString();
                    tel2.Text = MyData.dtmain.Rows[0]["c3"].ToString();

                }
                catch { }
            }
        }

        private void txtOhdelikveren_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    MyData.selectCommand("baza.accdb", "Select * from etibarnameneqliyyat where c3 Like " + "'%" + txtOhdelikveren.Text + "%'");
                    MyData.dtmain = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmain);

                    DateTime dt = DateTime.Now;

                    txtOhdelikveren.Text = MyData.dtmain.Rows[0]["c3"].ToString();
                    txtMuqNOMRESI.Text = MyData.dtmain.Rows[0]["c4"].ToString();
                    txtOhdelikNo.Text = txtMuqNOMRESI.Text.Substring(0, 1) + "-01/" + dt.Year.ToString().Substring(2, 2) + "-" + MyData.dtmain.Rows[0]["c4"].ToString().Substring(2, MyData.dtmain.Rows[0]["c4"].ToString().Length - 2);
                    txtobyekt.Text = "1 ədəd " + MyData.dtmain.Rows[0]["c2"].ToString() + " markalı avtomobil";
                    //txtavadanliq.Text = MyData.dtmain.Rows[0][5].ToString();
                    //txtfaiz.Text = MyData.dtmain.Rows[0][8].ToString();
                    //txtmuqTARIX.Text = MyData.dtmain.Rows[0][19].ToString();
                    //txtmuddet.Text = MyData.dtmain.Rows[0][6].ToString();
                    txtnomre.Text = MyData.dtmain.Rows[0]["c1"].ToString();
                    txtsehadetname.Text = MyData.dtmain.Rows[0]["c12"].ToString();
                    txtsassi.Text = MyData.dtmain.Rows[0]["c10"].ToString();
                    txtban.Text = MyData.dtmain.Rows[0]["c8"].ToString();
                    txtmuherrik.Text = MyData.dtmain.Rows[0]["c9"].ToString();
                    txtrengi.Text = MyData.dtmain.Rows[0]["c5"].ToString();
                    txtburaxilis.Text = MyData.dtmain.Rows[0]["c6"].ToString();

                    MyData.selectCommand("baza.accdb", "Select * from muqavilelayihe where [Lizinq alan] Like " + "'%" + txtOhdelikveren.Text + "%'");
                    MyData.dtmain = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmain);

                    txtavadanliqdeyer.Text = MyData.dtmain.Rows[0]["Avadanlığın dəyəri"].ToString();
                    txtfaiz.Text = MyData.dtmain.Rows[0]["% dərəcəsi"].ToString();
                    txtmuqTARIX.Text = MyData.dtmain.Rows[0]["Tarix"].ToString();
                    txtmuddet.Text = MyData.dtmain.Rows[0]["Lizinqin müddəti"].ToString();
                }

            }
            catch { };
        }

        private void txtMuqNOMRESI_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    MyData.selectCommand("baza.accdb", "Select * from etibarnameneqliyyat where c4 Like " + "'%" + txtMuqNOMRESI.Text + "%'");
                    MyData.dtmain = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmain);

                    DateTime dt = txthazirkiTARIX.Value.Date;

                    txtOhdelikveren.Text = MyData.dtmain.Rows[0]["c3"].ToString();
                    txtMuqNOMRESI.Text = MyData.dtmain.Rows[0]["c4"].ToString();
                    txtOhdelikNo.Text = txtMuqNOMRESI.Text.Substring(0, 1) + "-01/" + dt.Year.ToString().Substring(2, 2) + "-" + MyData.dtmain.Rows[0]["c4"].ToString().Substring(2, MyData.dtmain.Rows[0]["c4"].ToString().Length - 2);
                    txtobyekt.Text = "1 ədəd " + MyData.dtmain.Rows[0]["c2"].ToString().Replace("minik","") + " markalı avtomobil";
                    //txtavadanliq.Text = MyData.dtmain.Rows[0][5].ToString();
                    //txtfaiz.Text = MyData.dtmain.Rows[0][8].ToString();
                    //txtmuqTARIX.Text = MyData.dtmain.Rows[0][19].ToString();
                    //txtmuddet.Text = MyData.dtmain.Rows[0][6].ToString();
                    txtnomre.Text = MyData.dtmain.Rows[0]["c1"].ToString();
                    txtsehadetname.Text = MyData.dtmain.Rows[0]["c12"].ToString();
                    txtsassi.Text = MyData.dtmain.Rows[0]["c10"].ToString();
                    txtban.Text = MyData.dtmain.Rows[0]["c8"].ToString();
                    txtmuherrik.Text = MyData.dtmain.Rows[0]["c9"].ToString();
                    txtrengi.Text = MyData.dtmain.Rows[0]["c5"].ToString();
                    txtburaxilis.Text = MyData.dtmain.Rows[0]["c6"].ToString();

                    MyData.selectCommand("baza.accdb", "Select * from muqavilelayihe where [Lizinq alan] Like '%" + txtOhdelikveren.Text + "%'");
                    MyData.dtmain = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmain);

                    txtavadanliqdeyer.Text = MyData.dtmain.Rows[0]["Avadanlığın dəyəri"].ToString();
                    txtfaiz.Text = MyData.dtmain.Rows[0]["% dərəcəsi"].ToString();
                    txtmuqTARIX.Text = MyData.dtmain.Rows[0]["Tarix"].ToString();
                    txtmuddet.Text = MyData.dtmain.Rows[0]["Lizinqin müddəti"].ToString();
                }

            }
            catch { };
        }

        private void txtavadanliq_TextChanged(object sender, EventArgs e)
        {
            reqemler2();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (label43.Visible == false) { label43.Visible = true; label44.Visible = true; tel1.Visible = true; tel2.Visible = true; label28.Visible = true; label29.Visible = true; txtUnvan.Visible = true; txtSeriya.Visible = true; txtSeriyaNomre.Visible = true; return; }
            PrintYekunSifaris();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            PrintQrafik();
        }

        private void button15_Click(object sender, EventArgs e)
        {if (MyCheck.davamYesNo("Qeyd olunan nəqliyyat vasitəsinə dair əvvəlki məlumatlar ('Lizinq Alan' və 'Müqavilə Nömrəsi') Bazadan silinəcək. Yeniləmək istəyirsiniz?")) return;

            try
            {
                MyData.updateCommand("baza.accdb", "UPDATE etibarnameneqliyyat SET "
                                                                                     + "c3 ='" + txtOhdelikQebulEden.Text + "',"
                                                                                     + "c4 ='" + txtOhdelikNo.Text + "'"
                                                                                     + " WHERE c8='" + txtban.Text + "' and c9='" + txtmuherrik.Text + "' and c10='" + txtsassi.Text + "'");

                MessageBox.Show("Məlumat yeniləndi");
            }
            catch { MessageBox.Show("Əməliyyat baş tutmadı."); }
        }

        private void Ohdelik_Verilmesi_Load(object sender, EventArgs e)
        {

        }
    }
}
