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
using System.Speech.Synthesis;
using Nsoft;

namespace Lizinq_Muqavile
{
    public partial class DYP : Form
    {
        public DYP()
        {
            InitializeComponent();
        }

        Excel.Application oXL;
        Excel._Workbook oWB;
        Excel._Worksheet oSheet;

        public string n, m;

        public void DYPemr()        //----------------
        {
            try { Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Desktop\\DYP balansa goturme " + txtnomre.Text); File.Copy("DYP balansa goturme\\DYP emr balansa goturme.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\DYP balansa goturme " + txtnomre.Text + "\\DYP emr - " + txtnomre.Text + ".doc", true); }
            catch { MessageBox.Show("'DYP emr balansa goturme.doc' tapılmadı."); }

            object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\DYP balansa goturme " + txtnomre.Text + "\\DYP emr - " + txtnomre.Text + ".doc";


            //Create word Application Object
            Word.Application word = new Word.Application();

            //Create word document Object
            Word.Document doc = null;

            //Create word Missing Object
            object missing = System.Type.Missing;

            object readOnly = false;
            object isVisible = false;
            // make visible Word application
            word.Visible = true;

            doc = word.Documents.Open(ref FileName);
            doc.Activate();

            string a = MyChange.TarixSozle(Convert.ToDateTime(dtmuqtarix.Value));
            string b = MyChange.TarixSozle(Convert.ToDateTime(dttarix.Value));

            MyChange.FindAndReplace(word, "0000", dttarix.Text.Substring(0, 2) + " " + b + " " + dttarix.Text.Substring(6, 4) + "- cı il");
            MyChange.FindAndReplace(word, "000", txtmuqavilenomre.Text);
            MyChange.FindAndReplace(word, "111", dtmuqtarix.Text.Substring(0, 2) + " " + a + " " + dtmuqtarix.Text.Substring(6, 4) + "- cı il");
            MyChange.FindAndReplace(word, "222", txtlizinqalan.Text);
            MyChange.FindAndReplace(word, "444", txtmarka.Text);
           
            doc.Save();
        }
        public void DYPmektub()     //----------------
        {
            try { Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Desktop\\DYP balansa goturme " + txtnomre.Text); File.Copy("DYP balansa goturme\\DYP mektub balansa goturme.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\DYP balansa goturme " + txtnomre.Text + "\\DYP mektub - " + txtnomre.Text + ".doc", true); }
            catch { MessageBox.Show("'DYP mektub balansa goturme.doc' tapılmadı."); }

            object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\DYP balansa goturme " + txtnomre.Text + "\\DYP mektub - " + txtnomre.Text + ".doc";


            //Create word Application Object
            Word.Application word = new Word.Application();

            //Create word document Object
            Word.Document doc = null;

            //Create word Missing Object
            object missing = System.Type.Missing;

            object readOnly = false;
            object isVisible = false;
            // make visible Word application
            word.Visible = true;

            doc = word.Documents.Open(ref FileName);
            doc.Activate();

            string a = MyChange.TarixSozle(Convert.ToDateTime(dtmuqtarix.Value));
            string b = MyChange.TarixSozle(Convert.ToDateTime(dttarix.Value));

            MyChange.FindAndReplace(word, "0000", dttarix.Text.Substring(0, 2) + " " + b + " " + dttarix.Text.Substring(6, 4) + "- cı il");
            MyChange.FindAndReplace(word, "000", txtmuqavilenomre.Text);
            MyChange.FindAndReplace(word, "111", dtmuqtarix.Text.Substring(0, 2) + " " + a + " " + dtmuqtarix.Text.Substring(6, 4) + "- cı il");
            MyChange.FindAndReplace(word, "222", txtlizinqalan.Text);
            MyChange.FindAndReplace(word, "222", txtlizinqalan.Text);
            MyChange.FindAndReplace(word, "444", txtmarka.Text);
            MyChange.FindAndReplace(word, "444", txtmarka.Text);
            MyChange.FindAndReplace(word, "555", txtmuddet.Text);

            doc.Save();
        }
        private void EMRcopy()      //----------------
        {
            try { Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Desktop\\DYP balansa goturme " + txtnomre.Text); File.Copy("DYP balansa goturme\\Emr.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\DYP balansa goturme " + txtnomre.Text + "\\Emr - " + txtnomre.Text + ".xlsx", true); }
            catch { MessageBox.Show("'Emr.xlsx' tapılmadı."); }

            string muqaviletarixi = "", hazirkiTarixi = "", TexVerilmeTarixi = "";

            muqaviletarixi = Convert.ToDateTime(dtmuqtarix.Value).Day + " " + MyChange.TarixSozle(Convert.ToDateTime(dtmuqtarix.Value)) + " " + Convert.ToDateTime(dtmuqtarix.Value).Year;

            hazirkiTarixi = Convert.ToDateTime(dttarix.Value).Day + " " + MyChange.TarixSozle(Convert.ToDateTime(dttarix.Value)) + " " + Convert.ToDateTime(dttarix.Value).Year;

            TexVerilmeTarixi = Convert.ToDateTime(dttexpassverilmetarix.Value).Day + " " + MyChange.TarixSozle(Convert.ToDateTime(dttexpassverilmetarix.Value)) + " " + Convert.ToDateTime(dttexpassverilmetarix.Value).Year;

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\DYP balansa goturme " + txtnomre.Text + "\\Emr - " + txtnomre.Text + ".xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

            try
            {
                oSheet.Cells[4, 2] = "Tarix: " + hazirkiTarixi + "-ci il";
                oSheet.Cells[6, 1] = "            " + muqaviletarixi + "-ci il tarixli " + txtmuqavilenomre.Text + " saylı Daxili Maliyyə Lizinq müqaviləsi üzrə " + txtlizinqalan.Text + "na verilmiş 1 ədəd " + txtmarka.Text + " markalı avtomobili  “AGLizinq” QSC-nin balansına götürülsün.";
                if (txtmarka.Text.Substring(txtmarka.Text.Length - 5, 5) == "minik") oSheet.Cells[6, 1] = "            " + muqaviletarixi + "-ci il tarixli " + txtmuqavilenomre.Text + " saylı Daxili Maliyyə Lizinq müqaviləsi üzrə " + txtlizinqalan.Text + "na verilmiş 1 ədəd " + txtmarka.Text.Substring(0, txtmarka.Text.Length - 5) + " markalı minik avtomobili  “AGLizinq” QSC-nin balansına götürülsün.";
                if (txtmarka.Text.Substring(txtmarka.Text.Length - 15, 15) == "yəhərli dartıcı") oSheet.Cells[6, 1] = "            " + muqaviletarixi + "-ci il tarixli " + txtmuqavilenomre.Text + " saylı Daxili Maliyyə Lizinq müqaviləsi üzrə " + txtlizinqalan.Text + "na verilmiş 1 ədəd " + txtmarka.Text.Substring(0, txtmarka.Text.Length - 15) + " markalı yəhərli dartıcı  “AGLizinq” QSC-nin balansına götürülsün.";
                if (txtmarka.Text.Substring(txtmarka.Text.Length - 6, 6) == " qoşqu") oSheet.Cells[6, 1] = "            " + muqaviletarixi + "-ci il tarixli " + txtmuqavilenomre.Text + " saylı Daxili Maliyyə Lizinq müqaviləsi üzrə " + txtlizinqalan.Text + "na verilmiş 1 ədəd " + txtmarka.Text.Substring(0, txtmarka.Text.Length - 5) + " markalı qoşqu  “AGLizinq” QSC-nin balansına götürülsün.";
                if (txtmarka.Text.Substring(txtmarka.Text.Length - 10, 10) == "yarımqoşqu") oSheet.Cells[6, 1] = "            " + muqaviletarixi + "-ci il tarixli " + txtmuqavilenomre.Text + " saylı Daxili Maliyyə Lizinq müqaviləsi üzrə " + txtlizinqalan.Text + "na verilmiş 1 ədəd " + txtmarka.Text.Substring(0, txtmarka.Text.Length - 10) + " markalı yarımqoşqu  “AGLizinq” QSC-nin balansına götürülsün.";
                oSheet.Cells[2, 1] = "ƏMR № 000/15"; oSheet.Cells[2, 1].Font.Color = Color.Blue;
            }
            catch { };

            oXL.Visible = true;
            try
            {
                oXL.DisplayAlerts = false;
                oWB.Save();
            }
            catch { }
            // oXL.Application.Quit();
            // oXL.Visible = false;
            // oSheet.PrintOut();
            // oWB.Close(SaveChanges: false);
            // oXL.Workbooks.Close();
        }
        private void MEKTUBcopy()   //----------------
        {
            try { Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Desktop\\DYP balansa goturme " + txtnomre.Text); File.Copy("DYP balansa goturme\\Mektub.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\DYP balansa goturme " + txtnomre.Text + "\\Mektub - " + txtnomre.Text + ".xlsx", true); }
            catch { MessageBox.Show("'Mektub.xlsx' tapılmadı."); }

            string muqaviletarixi = "", hazirkiTarixi = "", TexVerilmeTarixi = "";

            muqaviletarixi = Convert.ToDateTime(dtmuqtarix.Value).Day + " " + MyChange.TarixSozle(Convert.ToDateTime(dtmuqtarix.Value)) + " " + Convert.ToDateTime(dtmuqtarix.Value).Year;

            hazirkiTarixi = Convert.ToDateTime(dttarix.Value).Day + " " + MyChange.TarixSozle(Convert.ToDateTime(dttarix.Value)) + " " + Convert.ToDateTime(dttarix.Value).Year;

            TexVerilmeTarixi = Convert.ToDateTime(dttexpassverilmetarix.Value).Day + " " + MyChange.TarixSozle(Convert.ToDateTime(dttexpassverilmetarix.Value)) + " " + Convert.ToDateTime(dttexpassverilmetarix.Value).Year;

            reqemler();

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\DYP balansa goturme " + txtnomre.Text + "\\Mektub - " + txtnomre.Text + ".xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

            try
            {
                oSheet.Cells[4, 1] = hazirkiTarixi + "-ci il";
                oSheet.Cells[8, 1] = "            Sizə bildiririk ki, “AGLizinq” Qapalı Səhmdar Cəmiyyəti ilə " + txtlizinqalan.Text + " arasında " + muqaviletarixi + "-ci il tarixdə imzalanmış " + txtmuqavilenomre.Text + " saylı daxili maliyyə lizinqi müqaviləsinə əsasən 1 ədəd " + txtmarka.Text + " markalı maşını " + txtlizinqalan.Text + "na " + txtmuddet.Text + " (" + m + ") ay müddətinə lizinqə verilmişdir." + Environment.NewLine + "           Yuxarıda qeyd olunanları nəzərə alaraq Sizdən xahiş edirik ki, 1 ədəd " + txtmarka.Text + " markalı maşını “AGLizinq” QSC-nin adına dövlət qeydiyyatına alasınız.";
                if (txtmarka.Text.Substring(txtmarka.Text.Length - 5, 5) == "minik") oSheet.Cells[8, 1] = "            Sizə bildiririk ki, “AGLizinq” Qapalı Səhmdar Cəmiyyəti ilə " + txtlizinqalan.Text + " arasında " + muqaviletarixi + "-ci il tarixdə imzalanmış " + txtmuqavilenomre.Text + " saylı daxili maliyyə lizinqi müqaviləsinə əsasən 1 ədəd " + txtmarka.Text.Substring(0, txtmarka.Text.Length - 5) + " markalı minik maşını " + txtlizinqalan.Text + "na " + txtmuddet.Text + " (" + m + ") ay müddətinə lizinqə verilmişdir." + Environment.NewLine + "           Yuxarıda qeyd olunanları nəzərə alaraq Sizdən xahiş edirik ki, 1 ədəd " + txtmarka.Text.Substring(0, txtmarka.Text.Length - 5) + " markalı minik maşınını “AGLizinq” QSC-nin adına dövlət qeydiyyatına alasınız.";
                if (txtmarka.Text.Substring(txtmarka.Text.Length - 15, 15) == "yəhərli dartıcı") oSheet.Cells[8, 1] = "            Sizə bildiririk ki, “AGLizinq” Qapalı Səhmdar Cəmiyyəti ilə " + txtlizinqalan.Text + " arasında " + muqaviletarixi + "-ci il tarixdə imzalanmış " + txtmuqavilenomre.Text + " saylı daxili maliyyə lizinqi müqaviləsinə əsasən 1 ədəd " + txtmarka.Text.Substring(0, txtmarka.Text.Length - 15) + " markalı yəhərli dartıcı " + txtlizinqalan.Text + "na " + txtmuddet.Text + " (" + m + ") ay müddətinə lizinqə verilmişdir." + Environment.NewLine + "           Yuxarıda qeyd olunanları nəzərə alaraq Sizdən xahiş edirik ki, 1 ədəd " + txtmarka.Text.Substring(0, txtmarka.Text.Length - 15) + " markalı yəhərli dartıcını “AGLizinq” QSC-nin adına dövlət qeydiyyatına alasınız.";
                if (txtmarka.Text.Substring(txtmarka.Text.Length - 6, 6) == " qoşqu") oSheet.Cells[8, 1] = "            Sizə bildiririk ki, “AGLizinq” Qapalı Səhmdar Cəmiyyəti ilə " + txtlizinqalan.Text + " arasında " + muqaviletarixi + "-ci il tarixdə imzalanmış " + txtmuqavilenomre.Text + " saylı daxili maliyyə lizinqi müqaviləsinə əsasən 1 ədəd " + txtmarka.Text.Substring(0, txtmarka.Text.Length - 5) + " markalı qoşqu " + txtlizinqalan.Text + "na " + txtmuddet.Text + " (" + m + ") ay müddətinə lizinqə verilmişdir." + Environment.NewLine + "           Yuxarıda qeyd olunanları nəzərə alaraq Sizdən xahiş edirik ki, 1 ədəd " + txtmarka.Text.Substring(0, txtmarka.Text.Length - 5) + " markalı qoşqunu “AGLizinq” QSC-nin adına dövlət qeydiyyatına alasınız.";
                if (txtmarka.Text.Substring(txtmarka.Text.Length - 10, 10) == "yarımqoşqu") oSheet.Cells[8, 1] = "            Sizə bildiririk ki, “AGLizinq” Qapalı Səhmdar Cəmiyyəti ilə " + txtlizinqalan.Text + " arasında " + muqaviletarixi + "-ci il tarixdə imzalanmış " + txtmuqavilenomre.Text + " saylı daxili maliyyə lizinqi müqaviləsinə əsasən 1 ədəd " + txtmarka.Text.Substring(0, txtmarka.Text.Length - 10) + " markalı yarımqoşqu " + txtlizinqalan.Text + "na " + txtmuddet.Text + " (" + m + ") ay müddətinə lizinqə verilmişdir." + Environment.NewLine + "           Yuxarıda qeyd olunanları nəzərə alaraq Sizdən xahiş edirik ki, 1 ədəd " + txtmarka.Text.Substring(0, txtmarka.Text.Length - 10) + " markalı yarımqoşqunu “AGLizinq” QSC-nin adına dövlət qeydiyyatına alasınız.";
                oSheet.Cells[2, 1] = "№ 000/15"; oSheet.Cells[2, 1].Font.Color = Color.Blue;
            }
            catch { };

            oXL.Visible = true;
            try
            {
                oXL.DisplayAlerts = false;
                oWB.Save();
            }
            catch { }
            // oXL.Application.Quit();
            // oXL.Visible = false;
            // oSheet.PrintOut();
            // oWB.Close(SaveChanges: false);
            // oXL.Workbooks.Close();
        }
        private void ERIZEcopy()    //----------------
        {
            try { Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Desktop\\DYP balansa goturme " + txtnomre.Text); File.Copy("DYP balansa goturme\\Erize.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\DYP balansa goturme " + txtnomre.Text + "\\Erize - " + txtnomre.Text + ".xlsx", true); }
            catch { MessageBox.Show("'Erize.xlsx' tapılmadı."); }

            int s;
            string k = "";

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\DYP balansa goturme " + txtnomre.Text + "\\Erize - " + txtnomre.Text + ".xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

            try
            {
                for (s = 0; s < txtmarka.Text.Length; s++)
                {
                    if (txtmarka.Text.Substring(s, 1) == " ") k = txtmarka.Text.Substring(1, s-2);

                }


                oSheet.Cells[20, 6] = k;
                oSheet.Cells[20, 1] = txttipi.Text;
                oSheet.Cells[22, 1] = txtbantipi.Text;
                oSheet.Cells[22, 6] = txtzavod.Text;
                oSheet.Cells[24, 1] = txtburaxilis.Text;
                oSheet.Cells[24, 6] = txtmuherrik.Text;
                oSheet.Cells[26, 1] = txtban.Text;
                oSheet.Cells[26, 6] = txtsassi.Text;
                oSheet.Cells[27, 1] = "9. Maksimum kütləsi:  " + txtmaxkutle.Text +" kq";
                oSheet.Cells[28, 6] = txtreng.Text;
                oSheet.Cells[29, 1] = "11. Yüksüz kütləsi:  " + txtyuksuzkutle.Text + " kq";
                oSheet.Cells[29, 4] = "M.İ.H. " + txtmih.Text + " sm3";
                oSheet.Cells[30, 6] = txtnomre.Text;
                oSheet.Cells[32, 1] = txtqeydiyyatsehadetname.Text;
                oSheet.Cells[32, 6] = txttranzit.Text;
                oSheet.Cells[34, 1] = txtetibaredilensexs.Text;
            }
            catch { };

            oXL.Visible = true;
            try
            {
                oXL.DisplayAlerts = false;
                oWB.Save();
            }
            catch { }
            // oXL.Application.Quit();
            // oXL.Visible = false;
            // oSheet.PrintOut();
            // oWB.Close(SaveChanges: false);
            // oXL.Workbooks.Close();
        }
        private void TTESLIMcopy()  //----------------
        {
            try { Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Desktop\\DYP balansa goturme " + txtnomre.Text); File.Copy("DYP balansa goturme\\Tehvil-Teslim.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\DYP balansa goturme " + txtnomre.Text + "\\Tehvil-Teslim - " + txtnomre.Text + ".xlsx", true); }
            catch { MessageBox.Show("'Tehvil-Teslim.xlsx' tapılmadı."); }


            int s;
            string k = "";

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\DYP balansa goturme " + txtnomre.Text + "\\Tehvil-Teslim - " + txtnomre.Text + ".xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;

            try
            {
                oSheet.Cells[11, 1] = txttipi.Text;
                oSheet.Cells[11, 7] = txtbantipi.Text;
                oSheet.Cells[11, 13] = txtburaxilis.Text;
                oSheet.Cells[11, 16] = txtzavod.Text;
                oSheet.Cells[11, 19] = txtmaxkutle.Text;
                oSheet.Cells[11, 22] = txtyuksuzkutle.Text;
                oSheet.Cells[11, 25] = txtqeydiyyatsehadetname.Text;

                for (s = 0; s < txtmarka.Text.Length; s++)
                {
                    if (txtmarka.Text.Substring(s, 1) == " ") k = txtmarka.Text.Substring(1, s - 2);

                }

                oSheet.Cells[18, 1] = k;
                oSheet.Cells[18, 7] = txtban.Text;
                oSheet.Cells[18, 12] = txtmuherrik.Text;
                oSheet.Cells[18, 15] = txtsassi.Text;
                oSheet.Cells[18, 20] = txtreng.Text;
                oSheet.Cells[18, 23] = txtnomre.Text;
                oSheet.Cells[18, 26] = txttranzit.Text;
                oSheet.Cells[49, 5] = txtSatici.Text;
            }
            catch { };

            oXL.Visible = true;
            try
            {
                oXL.DisplayAlerts = false;
                oWB.Save();
            }
            catch { }
            // oXL.Application.Quit();
            // oXL.Visible = false;
            // oSheet.PrintOut();
            // oWB.Close(SaveChanges: false);
            // oXL.Workbooks.Close();
        }
        private void ASATQIcopy()   //----------------
        {
            try { Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Desktop\\DYP balansa goturme " + txtnomre.Text); File.Copy("DYP balansa goturme\\Alqı-Satqı.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\DYP balansa goturme " + txtnomre.Text + "\\Alqı-Satqı - " + txtnomre.Text + ".xlsx", true); }
            catch { MessageBox.Show("'Alqı-Satqı.xlsx' tapılmadı."); }

            int s;
            string k = "", muqaviletarixi = "", hazirkiTarixi = "", TexVerilmeTarixi = "";

            muqaviletarixi = Convert.ToDateTime(dtmuqtarix.Value).Day + " " + MyChange.TarixSozle(Convert.ToDateTime(dtmuqtarix.Value)) + " " + Convert.ToDateTime(dtmuqtarix.Value).Year;

            hazirkiTarixi = Convert.ToDateTime(dttarix.Value).Day + " " + MyChange.TarixSozle(Convert.ToDateTime(dttarix.Value)) + " " + Convert.ToDateTime(dttarix.Value).Year;

            TexVerilmeTarixi = Convert.ToDateTime(dttexpassverilmetarix.Value).Day + " " + MyChange.TarixSozle(Convert.ToDateTime(dttexpassverilmetarix.Value)) + " " + Convert.ToDateTime(dttexpassverilmetarix.Value).Year;

            reqemler2();

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\DYP balansa goturme " + txtnomre.Text + "\\Alqı-Satqı - " + txtnomre.Text + ".xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

            try
            {
                oSheet.Cells[12, 9] = hazirkiTarixi.Substring(3,hazirkiTarixi.Length-3) + "-ci il";
                oSheet.Cells[14, 2] = txtSatici.Text + ", " + txtsexsiyyetsatici.Text + ", " + txtsexsiyyetverilmetarix.Text + " " + txtverenorqan.Text;
                oSheet.Cells[16, 1] = txtunvansatici.Text;

                for (s = 0; s < txtmarka.Text.Length; s++)
                {
                    if (txtmarka.Text.Substring(s, 1) == " ") k = txtmarka.Text.Substring(1, s - 2);

                }
                oSheet.Cells[27, 7] = k;
                oSheet.Cells[28, 3] = txtnomre.Text;
                oSheet.Cells[28, 6] = txtburaxilis.Text;
                oSheet.Cells[28, 9] = txtmuherrik.Text;
                oSheet.Cells[29, 2] = txtsassi.Text;
                oSheet.Cells[29, 7] = txtban.Text;
                oSheet.Cells[30, 1] = "< " + Convert.ToDateTime(dttexpassverilmetarix.Value).Day + " >";
                oSheet.Cells[30, 2] = MyChange.TarixSozle(Convert.ToDateTime(dttexpassverilmetarix.Value));
                oSheet.Cells[30, 4] = Convert.ToDateTime(dttexpassverilmetarix.Value).Year + "-ci ildə DYP tərəfindən verilmiş";
                oSheet.Cells[34, 1] = txtqeydiyyatsehadetname.Text.Substring(0,2);
                oSheet.Cells[34, 4] = "'" + txtqeydiyyatsehadetname.Text.Substring(txtqeydiyyatsehadetname.Text.Length - 6, 6);
                oSheet.Cells[36, 3] = txtalqisatqimebleg.Text + " (" + m + ")";
                oSheet.Cells[48, 2] = hazirkiTarixi.Substring(3, hazirkiTarixi.Length - 3);
                oSheet.Cells[52, 6] = hazirkiTarixi.Substring(3, hazirkiTarixi.Length - 3);
            }
            catch { };

            oXL.Visible = true;
            try
            {
                oXL.DisplayAlerts = false;
                oWB.Save();
            }
            catch { }
            // oXL.Application.Quit();
            // oXL.Visible = false;
            // oSheet.PrintOut();
            // oWB.Close(SaveChanges: false);
            // oXL.Workbooks.Close();
        }
        private void ERIZEprint()   //----------------
        {
            try { Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Desktop\\DYP balansa goturme " + txtnomre.Text); File.Copy("DYP balansa goturme\\Erize.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\DYP balansa goturme " + txtnomre.Text + "\\Erize - " + txtnomre.Text + ".xlsx", true); }
            catch { MessageBox.Show("'\\192.168.10.5\\Common\\AGLizinq\\Erize.xlsx' tapılmadı."); }

            int s;
            string k = "";
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\DYP balansa goturme " + txtnomre.Text + "\\Erize - " + txtnomre.Text + ".xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

            try
            {
                for (s = 0; s < txtmarka.Text.Length; s++)
                {
                    if (txtmarka.Text.Substring(s, 1) == " ") k = txtmarka.Text.Substring(1, s - 2);

                }


                oSheet.Cells[20, 6] = k;
                oSheet.Cells[20, 1] = txttipi.Text;
                oSheet.Cells[22, 1] = txtbantipi.Text;
                oSheet.Cells[22, 6] = txtzavod.Text;
                oSheet.Cells[24, 1] = txtburaxilis.Text;
                oSheet.Cells[24, 6] = txtmuherrik.Text;
                oSheet.Cells[26, 1] = txtban.Text;
                oSheet.Cells[26, 6] = txtsassi.Text;
                oSheet.Cells[27, 1] = "9. Maksimum kütləsi:  " + txtmaxkutle.Text + " kq";
                oSheet.Cells[28, 6] = txtreng.Text;
                oSheet.Cells[29, 1] = "11. Yüksüz kütləsi:  " + txtyuksuzkutle.Text + " kq";
                oSheet.Cells[29, 4] = "M.İ.H. " + txtmih.Text + " sm3";
                oSheet.Cells[30, 6] = txtnomre.Text;
                oSheet.Cells[32, 1] = txtqeydiyyatsehadetname.Text;
                oSheet.Cells[32, 6] = txttranzit.Text;
                oSheet.Cells[34, 1] = txtetibaredilensexs.Text;
            }
            catch { };

            oXL.Visible = false;
            try
            {
                oXL.DisplayAlerts = false;
                oWB.Save();
            }
            catch { }
            
            oXL.Visible = false;
            oSheet.PrintOut();
            oWB.Close(SaveChanges: true);
            oXL.Workbooks.Close();
            oXL.Application.Quit();
        }
        private void TTESLIMprint() //----------------
        {
            try { Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Desktop\\DYP balansa goturme " + txtnomre.Text); File.Copy("DYP balansa goturme\\Tehvil-Teslim.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\DYP balansa goturme " + txtnomre.Text + "\\Tehvil-Teslim - " + txtnomre.Text + ".xlsx", true); }
            catch { MessageBox.Show("'\\192.168.10.5\\Common\\AGLizinq\\Tehvil-Teslim.xlsx' tapılmadı."); }

            int s;
            string k = "";
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\DYP balansa goturme " + txtnomre.Text + "\\Tehvil-Teslim - " + txtnomre.Text + ".xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;

            try
            {
                oSheet.Cells[11, 1] = txttipi.Text;
                oSheet.Cells[11, 7] = txtbantipi.Text;
                oSheet.Cells[11, 13] = txtburaxilis.Text;
                oSheet.Cells[11, 16] = txtzavod.Text;
                oSheet.Cells[11, 19] = txtmaxkutle.Text;
                oSheet.Cells[11, 22] = txtyuksuzkutle.Text;
                oSheet.Cells[11, 25] = txtqeydiyyatsehadetname.Text;

                for (s = 0; s < txtmarka.Text.Length; s++)
                {
                    if (txtmarka.Text.Substring(s, 1) == " ") k = txtmarka.Text.Substring(1, s - 2);

                }

                oSheet.Cells[18, 1] = k;
                oSheet.Cells[18, 7] = txtban.Text;
                oSheet.Cells[18, 12] = txtmuherrik.Text;
                oSheet.Cells[18, 15] = txtsassi.Text;
                oSheet.Cells[18, 20] = txtreng.Text;
                oSheet.Cells[18, 23] = txtnomre.Text;
                oSheet.Cells[18, 26] = txttranzit.Text;
                oSheet.Cells[49, 5] = txtSatici.Text;
            }
            catch { };

            oXL.Visible = false;
            try
            {
                oXL.DisplayAlerts = false;
                oWB.Save();
            }
            catch { }
            
            oXL.Visible = false;
            oSheet.PrintOut(1,1);
            MessageBox.Show("Təhvil Təslim aktı Vərəqini tərsinə çevirib yenidən printerə yerləşdirin və 'OK' düyməsinə basın...");
            oSheet.PrintOut(2, 2);
            oWB.Close(SaveChanges: true);
            oXL.Workbooks.Close();
            oXL.Application.Quit();
        }
        private void ASATQIprint()  //----------------
        {
            try { Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Desktop\\DYP balansa goturme " + txtnomre.Text); File.Copy("DYP balansa goturme\\Alqı-Satqı.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\DYP balansa goturme " + txtnomre.Text + "\\Alqı-Satqı - " + txtnomre.Text + ".xlsx", true); }
            catch { MessageBox.Show("'Alqı-Satqı.xlsx' tapılmadı."); }

            int s;
            string k = "", muqaviletarixi = "", hazirkiTarixi = "", TexVerilmeTarixi = "";

            muqaviletarixi = Convert.ToDateTime(dtmuqtarix.Value).Day + " " + MyChange.TarixSozle(Convert.ToDateTime(dtmuqtarix.Value)) + " " + Convert.ToDateTime(dtmuqtarix.Value).Year;

            hazirkiTarixi = Convert.ToDateTime(dttarix.Value).Day + " " + MyChange.TarixSozle(Convert.ToDateTime(dttarix.Value)) + " " + Convert.ToDateTime(dttarix.Value).Year;

            TexVerilmeTarixi = Convert.ToDateTime(dttexpassverilmetarix.Value).Day + " " + MyChange.TarixSozle(Convert.ToDateTime(dttexpassverilmetarix.Value)) + " " + Convert.ToDateTime(dttexpassverilmetarix.Value).Year;

            reqemler2();

            

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\DYP balansa goturme " + txtnomre.Text + "\\Alqı-Satqı - " + txtnomre.Text + ".xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

            try
            {
                oSheet.Cells[12, 9] = hazirkiTarixi.Substring(3, hazirkiTarixi.Length - 3) + "-ci il";
                oSheet.Cells[14, 2] = txtSatici.Text + ", " + txtsexsiyyetsatici.Text + ", " + txtsexsiyyetverilmetarix.Text + " " + txtverenorqan.Text;
                oSheet.Cells[16, 1] = txtunvansatici.Text;

                for (s = 0; s < txtmarka.Text.Length; s++)
                {
                    if (txtmarka.Text.Substring(s, 1) == " ") k = txtmarka.Text.Substring(1, s - 2);

                }
                oSheet.Cells[27, 7] = k;
                oSheet.Cells[28, 3] = txtnomre.Text;
                oSheet.Cells[28, 6] = txtburaxilis.Text;
                oSheet.Cells[28, 9] = txtmuherrik.Text;
                oSheet.Cells[29, 2] = txtsassi.Text;
                oSheet.Cells[29, 7] = txtban.Text;
                oSheet.Cells[30, 1] = "< " + Convert.ToDateTime(dttexpassverilmetarix.Value).Day + " >";
                oSheet.Cells[30, 2] = MyChange.TarixSozle(Convert.ToDateTime(dttexpassverilmetarix.Value));
                oSheet.Cells[30, 4] = Convert.ToDateTime(dttexpassverilmetarix.Value).Year + "-cü ildə DYP tərəfindən verilmiş";
                oSheet.Cells[34, 1] = txtqeydiyyatsehadetname.Text.Substring(0, 2);
                oSheet.Cells[34, 4] = "'" + txtqeydiyyatsehadetname.Text.Substring(txtqeydiyyatsehadetname.Text.Length - 6, 6);
                oSheet.Cells[36, 3] = txtalqisatqimebleg.Text + " (" + m + ")";
                oSheet.Cells[48, 2] = hazirkiTarixi.Substring(3, hazirkiTarixi.Length - 3);
                oSheet.Cells[52, 6] = hazirkiTarixi.Substring(3, hazirkiTarixi.Length - 3);
            }
            catch { };

            try
            {
                oXL.DisplayAlerts = false;
                oWB.Save();
            }
            catch { }
            
            oXL.Visible = false;
            oSheet.PrintOut();
            oWB.Close(SaveChanges: true);
            oXL.Workbooks.Close();
            oXL.Application.Quit();
        }
        private void reqemler()       //------reqem yazi ile---------------------------------------------------------------
        {
            try { m = MyChange.ReqemToMetn(Convert.ToInt32(txtmuddet.Text)); }catch{ }
        }
        private void reqemler2()       //------reqem yazi ile---------------------------------------------------------------
        {
            try { m = MyChange.ReqemToMetn(Convert.ToInt32(txtalqisatqimebleg.Text)); } catch { }
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
           try
           {
               if (e.KeyCode == Keys.Enter)
               {
                MyData.selectCommand("baza.accdb","Select * from etibarnameneqliyyat where c1 Like '%" + txtnomre.Text + "%'");
                MyData.dtmain = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmain);

                txtnomre.Text = MyData.dtmain.Rows[0]["c1"].ToString(); 
                txtmarka.Text = MyData.dtmain.Rows[0]["c2"].ToString(); 
                txtburaxilis.Text = MyData.dtmain.Rows[0]["c6"].ToString();
                txtban.Text = MyData.dtmain.Rows[0]["c8"].ToString();
                txtqeydiyyatsehadetname.Text = MyData.dtmain.Rows[0]["c12"].ToString();
                txtsassi.Text = MyData.dtmain.Rows[0]["c10"].ToString();
                txtreng.Text = MyData.dtmain.Rows[0]["c5"].ToString();
                txtmuqavilenomre.Text = MyData.dtmain.Rows[0]["c4"].ToString();
                txtmuherrik.Text = MyData.dtmain.Rows[0]["c9"].ToString();
                txtlizinqalan.Text = MyData.dtmain.Rows[0]["c3"].ToString();
                dttexpassverilmetarix.Text = MyData.dtmain.Rows[0]["c7"].ToString();

                }
            }
            catch { };
        }

        private void textBoxSatici_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    MyData.selectCommand("baza.accdb", "Select * from muqavilesaticirekvizit WHERE Satıcı Like '%" + txtSatici.Text + "%'");
                    MyData.dtmain = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmain);

                    txtSatici.Text = MyData.dtmain.Rows[0]["Satıcı"].ToString(); // AD familiya
                    txtsexsiyyetsatici.Text = MyData.dtmain.Rows[0]["Şəxsiyyət vəsiqəsi"].ToString(); //Ş\V
                    txtunvansatici.Text = MyData.dtmain.Rows[0]["Ş/V qeydiyyat ünvanı"].ToString();
                    txtsexsiyyetverilmetarix.Text = MyData.dtmain.Rows[0]["Ş/V verilmə tarixi"].ToString();
                    txtverenorqan.Text = MyData.dtmain.Rows[0]["Ş/V verən orqan"].ToString();
                }
            }
            catch { };
        }

        private void label16_Click(object sender, EventArgs e)
        {
            if (txtunvansatici.Enabled == false) { txtunvansatici.Enabled = true; return; }
            txtunvansatici.Enabled = false;
        }

        private void label13_Click(object sender, EventArgs e)
        {
            if (txtsexsiyyetsatici.Enabled == false) { txtsexsiyyetsatici.Enabled = true; return; }
            txtsexsiyyetsatici.Enabled = false;

        }

        private void label14_Click(object sender, EventArgs e)
        {
            if (txtsexsiyyetverilmetarix.Enabled == false) { txtsexsiyyetverilmetarix.Enabled = true; return; }
            txtsexsiyyetverilmetarix.Enabled = false;
        }

        private void label15_Click(object sender, EventArgs e)
        {
            if (txtverenorqan.Enabled == false) { txtverenorqan.Enabled = true; return; }
            txtverenorqan.Enabled = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            EMRcopy();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            MEKTUBcopy();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            ERIZEcopy();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            TTESLIMcopy();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            ASATQIcopy();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            ERIZEprint();
            ASATQIprint();
            TTESLIMprint();

            EMRcopy();
            MEKTUBcopy();
        }

        private void çıxışToolStripMenuItem_Click(object sender, EventArgs e)
        {
            base.Close();
        }

        private void köməkToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Kömək üçün 'F1' düyməsindən istifadə edin.", "Kömək");
        }

        private void DYP_Load(object sender, EventArgs e)
        {
            DateTime dt = DateTime.Now;
            dtmuqtarix.Text = dt.ToShortDateString();    // Muqavile tarixi
            dttarix.Text = dt.ToShortDateString();    // Hazirki tarix
        }

        private void label116_Click(object sender, EventArgs e)
        {
            if (txtburaxilis.Enabled == false) { txtburaxilis.Enabled = true; return; }
            txtburaxilis.Enabled = false;
        }

        private void label120_Click(object sender, EventArgs e)
        {
            if (txtban.Enabled == false) { txtban.Enabled = true; return; }
            txtban.Enabled = false;
        }

        private void label2_Click(object sender, EventArgs e)
        {
            if (txtmarka.Enabled == false) { txtmarka.Enabled = true; return; }
            txtmarka.Enabled = false;
        }

        private void label119_Click(object sender, EventArgs e)
        {
            if (txtmuherrik.Enabled == false) { txtmuherrik.Enabled = true; return; }
            txtmuherrik.Enabled = false;
        }

        private void label122_Click(object sender, EventArgs e)
        {
            if (txtreng.Enabled == false) { txtreng.Enabled = true; return; }
            txtreng.Enabled = false;
        }

        private void label74_Click(object sender, EventArgs e)
        {
            if (txtmuqavilenomre.Enabled == false) { txtmuqavilenomre.Enabled = true; return; }
            txtmuqavilenomre.Enabled = false;
        }

        private void label21_Click(object sender, EventArgs e)
        {
            if (txtlizinqalan.Enabled == false) { txtlizinqalan.Enabled = true; return; }
            txtlizinqalan.Enabled = false;
        }

        private void label118_Click(object sender, EventArgs e)
        {
            if (txtsassi.Enabled == false) { txtsassi.Enabled = true; return; }
            txtsassi.Enabled = false;
        }

        private void button13_Click(object sender, EventArgs e)
        {
            DYPemr();
            DYPmektub();
        }
    }
}
 