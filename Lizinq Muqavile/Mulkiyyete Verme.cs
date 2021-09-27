using Nsoft;
using System;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Data;


namespace Lizinq_Muqavile
{
    public partial class Mulkiyyete_Verme : Form
    {

        Excel.Application oXL;
        Excel._Workbook oWB;
        Excel._Worksheet oSheet;
        Excel.Application oXL2;
        Excel._Workbook oWB2;
        Excel._Worksheet oSheet2;

        public Mulkiyyete_Verme()
        {
            InitializeComponent();
        }

        public int emeliyyatUcun = 0;

        private void AutoDuzelis()
        {
            try
            {
                if (txtmarka.Text.Substring(txtmarka.Text.Length - 5, 5) == "minik") { txttip.Text = "Minik"; txtmarka.Text = txtmarka.Text.Substring(0, txtmarka.Text.Length - 6); }
                else if (txtmarka.Text.Substring(txtmarka.Text.Length - 3, 3) == "yük") { txttip.Text = "Yük"; txtmarka.Text = txtmarka.Text.Substring(0, txtmarka.Text.Length - 4); }
                else if (txtmarka.Text.Substring(txtmarka.Text.Length - 10, 10) == "yarımqoşqu") { txttip.Text = "Yarımqoşqu"; txtmarka.Text = txtmarka.Text.Substring(0, txtmarka.Text.Length - 11); }
                else if (txtmarka.Text.Substring(txtmarka.Text.Length - 19, 19) == "tırtıllı ekskavator") { txttip.Text = "Ekskavator"; txtmodel.Text = "Tırtıllı"; txtmarka.Text = txtmarka.Text.Substring(0, txtmarka.Text.Length - 20); }
                else if (txtmarka.Text.Substring(txtmarka.Text.Length - 10, 10) == "ekskavator") { txttip.Text = "Ekskavator"; txtmarka.Text = txtmarka.Text.Substring(0, txtmarka.Text.Length - 11); }
                else if (txtmarka.Text.Substring(txtmarka.Text.Length - 15, 15) == "yəhərli dartıcı") { txttip.Text = "Yük"; txtmodel.Text = "Yəhərli dartıcı"; txtmarka.Text = txtmarka.Text.Substring(0, txtmarka.Text.Length - 16); }
                else if (txtmarka.Text.Substring(txtmarka.Text.Length - 7, 7) == "avtobus") { txttip.Text = "avtobus"; txtmarka.Text = txtmarka.Text.Substring(0, txtmarka.Text.Length - 8); }
            }
            catch { }

            try
            {
                if (txtmarka2.Text.Substring(txtmarka2.Text.Length - 5, 5) == "minik") { txttip2.Text = "Minik"; txtmarka2.Text = txtmarka2.Text.Substring(0, txtmarka2.Text.Length - 6); }
                else if (txtmarka2.Text.Substring(txtmarka2.Text.Length - 3, 3) == "yük") { txttip2.Text = "Yük"; txtmarka2.Text = txtmarka2.Text.Substring(0, txtmarka2.Text.Length - 4); }
                else if (txtmarka2.Text.Substring(txtmarka2.Text.Length - 10, 10) == "yarımqoşqu") { txttip2.Text = "Yarımqoşqu"; txtmarka2.Text = txtmarka2.Text.Substring(0, txtmarka2.Text.Length - 11); }
                else if (txtmarka2.Text.Substring(txtmarka2.Text.Length - 19, 19) == "tırtıllı ekskavator") { txttip2.Text = "Ekskavator"; txtmodel2.Text = "Tırtıllı"; txtmarka2.Text = txtmarka2.Text.Substring(0, txtmarka2.Text.Length - 20); }
                else if (txtmarka2.Text.Substring(txtmarka2.Text.Length - 10, 10) == "ekskavator") { txttip2.Text = "Ekskavator"; txtmarka2.Text = txtmarka2.Text.Substring(0, txtmarka2.Text.Length - 11); }
                else if (txtmarka2.Text.Substring(txtmarka2.Text.Length - 15, 15) == "yəhərli dartıcı") { txttip2.Text = "Yük"; txtmodel.Text = "Yəhərli dartıcı"; txtmarka2.Text = txtmarka2.Text.Substring(0, txtmarka2.Text.Length - 16); }
                else if (txtmarka2.Text.Substring(txtmarka2.Text.Length - 7, 7) == "avtobus") { txttip2.Text = "Avtobus"; txtmarka2.Text = txtmarka2.Text.Substring(0, txtmarka2.Text.Length - 8); }

            }
            catch { }
        }

        private void MMXrefresh()
        {
            try
            {
                MyData.selectCommand("baza.accdb", "SELECT * FROM MMX WHERE a2 like '" + txtnomre.Text + "' and a9 like 'Xeyr'");
                MyData.dtmainMMX = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainMMX);
            }
            catch { }

            int tt = 0;
            try
            {
                for (int t = 0; t < MyData.dtmainMMX.Rows.Count; t++)
                {
                    tt += Convert.ToInt32(MyData.dtmainMMX.Rows[t]["a10"]);
                }
            }
            catch { }

            if (MyData.dtmainMMX.Rows.Count > 0)
            {
                btMMX2.Text = MyData.dtmainMMX.Rows.Count.ToString() + " ƏDƏD MMX MÖVCUDDUR (" + tt.ToString() +" AZN)";
                btMMX2.ForeColor = Color.White;
                btMMX2.BackColor = Color.Red;
                btMMX2.FlatAppearance.BorderColor = Color.Red;
            }
            else
            {
                btMMX2.Text = "MMX YOXDUR.";
                btMMX2.ForeColor = Color.Green;
                btMMX2.BackColor = Color.Silver;
                btMMX2.FlatAppearance.BorderColor = Color.Gray;
            }
        }

        public void WordDoc()
        {
            try { Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text); File.Copy("Mulkiyyete verme\\Mulkiyyete verme.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text + "\\Mulkiyyete verme - " + txtnomre.Text + ".doc", true); }
            catch { MessageBox.Show("'MulkiyyeteVerme.doc' tapılmadı."); }

            if (dttarix.Text == dttarix2.Text)
            {
                if (!MyCheck.davamYesNo("Müqavilənin tarixinin düzgünlüyündən əminsinizmi?")) return;
            }

            object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text + "\\Mulkiyyete verme - " + txtnomre.Text + ".doc";


            Word.Application word = new Word.Application();
            Word.Document doc = null;
            object missing = System.Type.Missing;
            object readOnly = false;
            object isVisible = false;
            word.Visible = true;

            doc = word.Documents.Open(ref FileName);
            doc.Activate();

            DateTime dt2 = dttarix2.Value.Date;
            DateTime dt = dttarix.Value.Date;

            string a = MyChange.TarixSozle(dt2);
            string b = MyChange.TarixSozle(dt);

          MyChange.FindAndReplace(word, "000", txtlayihe.Text);
          MyChange.FindAndReplace(word, "000", txtlayihe.Text);
          MyChange.FindAndReplace(word, "000", txtlayihe.Text);
          MyChange.FindAndReplace(word, "111", dt2.Day + " " + a + " " + dt2.Year + "- ci il");
          MyChange.FindAndReplace(word, "111", dt2.Day + " " + a + " " + dt2.Year + "- ci il");
          MyChange.FindAndReplace(word, "111", dt2.Day + " " + a + " " + dt2.Year + "- ci il");
          MyChange.FindAndReplace(word, "111", dt2.Day + " " + a + " " + dt2.Year + "- ci il");
          MyChange.FindAndReplace(word, "222", dt.Day + " " + b + " " + dt.Year + "- ci il");
          MyChange.FindAndReplace(word, "222", dt.Day + " " + b + " " + dt.Year + "- ci il");
          MyChange.FindAndReplace(word, "333", txtlizinqalan.Text);
          MyChange.FindAndReplace(word, "333", txtlizinqalan.Text);
          MyChange.FindAndReplace(word, "333", txtlizinqalan.Text);
            
          MyChange.FindAndReplace(word, "444", txtmarka.Text + " " + txttip.Text);
          MyChange.FindAndReplace(word, "555", txtnomre.Text);
          MyChange.FindAndReplace(word, "666", txtsehadetname.Text);
          MyChange.FindAndReplace(word, "6666666", txtzavodnomresi.Text);
          MyChange.FindAndReplace(word, "777", txtsassi.Text);
          MyChange.FindAndReplace(word, "888", txtban.Text);
          MyChange.FindAndReplace(word, "999", txtmuherrik.Text);
          MyChange.FindAndReplace(word, "1111", txtrengi.Text);
          MyChange.FindAndReplace(word, "2222", txtburaxilis.Text);

            if (cb1.Checked == true)
            {
              MyChange.FindAndReplace(word, "C444", txtmarka2.Text + " " + txttip2.Text);
              MyChange.FindAndReplace(word, "C555", "Qeydiyyat nişanı – " + txtnomre2.Text);
              MyChange.FindAndReplace(word, "C666", "Qeydiyyat şəhadətnaməsi – " + txtsehadetname2.Text);
              MyChange.FindAndReplace(word, "C6666666", "Zavod - " + txtzavodnomresi2.Text);
              MyChange.FindAndReplace(word, "C777", "Şassi – " + txtsassi2.Text);
              MyChange.FindAndReplace(word, "C888", "Ban – " + txtban2.Text);
              MyChange.FindAndReplace(word, "C999", "Mühərrik – " + txtmuherrik2.Text);
              MyChange.FindAndReplace(word, "C1111", "Rəngi – " + txtrengi2.Text);
              MyChange.FindAndReplace(word, "C2222", "Buraxılış ili - " + txtburaxilis2.Text);
            }
            else
            {
              MyChange.FindAndReplace(word, "C444", "");
              MyChange.FindAndReplace(word, "C555", "");
              MyChange.FindAndReplace(word, "C666", "");
              MyChange.FindAndReplace(word, "C6666666", "");
              MyChange.FindAndReplace(word, "C777", "");
              MyChange.FindAndReplace(word, "C888", "");
              MyChange.FindAndReplace(word, "C999", "");
              MyChange.FindAndReplace(word, "C1111", "");
              MyChange.FindAndReplace(word, "C2222", "");
            }

          MyChange.FindAndReplace(word, "3333", cbmuqavileFormasi.Text);
          MyChange.FindAndReplace(word, "3333", cbmuqavileFormasi.Text);
          MyChange.FindAndReplace(word, "3333", cbmuqavileFormasi.Text);

            try
            {
                if (comboBox1.Text == "«AGBank» ASC")
                {
                  MyChange.FindAndReplace(word, "111111", "«AGBank» ASC-nin İdarə heyətinin sədri Ə.S Cəlilov (17.08.2017 tarixli etibarnaməyə əsasən) «AGLizinq» QSC-nin Baş Meneceri R.İ. Musayev şəxsində");
                  MyChange.FindAndReplace(word, "111112", comboBox1.Text);
                  MyChange.FindAndReplace(word, "111112", comboBox1.Text);
                  MyChange.FindAndReplace(word, "111112", comboBox1.Text);
                }
                else 
                {
                  MyChange.FindAndReplace(word, "111111", "Baş Menecer: Musayev Rəşad İslam oğlu şəxsində");
                  MyChange.FindAndReplace(word, "111112", comboBox1.Text);
                  MyChange.FindAndReplace(word, "111112", comboBox1.Text);
                  MyChange.FindAndReplace(word, "111112", comboBox1.Text);
                }

                
            }
            catch { }

            doc.Save();

            emeliyyatUcun = emeliyyatUcun + 1;

            if (emeliyyatUcun == 2)
            {
                MyData.insertCommand("baza.accdb", "insert into Emeliyyatlar (a1, a2, a3) Values ('" + dt.Date + "', 'Mülkiyyətə vermə - " + txtnomre.Text + " - " + txtlizinqalan.Text + " - " + txtmarka.Text + "','" + Environment.MachineName + "')");
            }
        }

        public void WordDocHuquqiSexs()
        {
            try { Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text); File.Copy("Mulkiyyete verme\\Mulkiyyete verme Huquqi Sexs.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text + "\\Mulkiyyete verme - " + txtnomre.Text + ".doc", true); }
            catch { MessageBox.Show("'\\192.168.10.5\\Common\\AGLizinq\\MulkiyyeteVerme.doc' tapılmadı."); }

            if (dttarix.Text == dttarix2.Text)
            {
                MyCheck.davamYesNo("Müqavilənin tarixinin düzgünlüyündən əminsinizmi?"); return;
            }

            object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text + "\\Mulkiyyete verme - " + txtnomre.Text + ".doc";

            Word.Application word = new Word.Application();
            Word.Document doc = null;
            object missing = System.Type.Missing;
            object readOnly = false;
            object isVisible = false;
            word.Visible = true;

            doc = word.Documents.Open(ref FileName);
            doc.Activate();

            DateTime dt2 = dttarix2.Value.Date;
            DateTime dt = dttarix.Value.Date;

            string a = MyChange.TarixSozle(dt2);
            string b = MyChange.TarixSozle(dt);

            MyChange.FindAndReplace(word, "000", txtlayihe.Text);
            MyChange.FindAndReplace(word, "000", txtlayihe.Text);
            MyChange.FindAndReplace(word, "000", txtlayihe.Text);
            MyChange.FindAndReplace(word, "111", dt2.Day + " " + a + " " + dt2.Year + "- ci il");
            MyChange.FindAndReplace(word, "111", dt2.Day + " " + a + " " + dt2.Year + "- ci il");
            MyChange.FindAndReplace(word, "111", dt2.Day + " " + a + " " + dt2.Year + "- ci il");
            MyChange.FindAndReplace(word, "111", dt2.Day + " " + a + " " + dt2.Year + "- ci il");
            MyChange.FindAndReplace(word, "222", dt.Day + " " + b + " " + dt.Year + "- ci il");
            MyChange.FindAndReplace(word, "222", dt.Day + " " + b + " " + dt.Year + "- ci il");
            MyChange.FindAndReplace(word, "333", txtlizinqalan.Text);
            MyChange.FindAndReplace(word, "333", txtlizinqalan.Text);
            MyChange.FindAndReplace(word, "333", txtlizinqalan.Text);

            MyChange.FindAndReplace(word, "444", txtmarka.Text);
            MyChange.FindAndReplace(word, "555", txtnomre.Text);
            MyChange.FindAndReplace(word, "666", txtsehadetname.Text);
            MyChange.FindAndReplace(word, "6666666", txtzavodnomresi.Text);
            MyChange.FindAndReplace(word, "777", txtsassi.Text);
            MyChange.FindAndReplace(word, "888", txtban.Text);
            MyChange.FindAndReplace(word, "999", txtmuherrik.Text);
            MyChange.FindAndReplace(word, "1111", txtrengi.Text);
            MyChange.FindAndReplace(word, "2222", txtburaxilis.Text);

            if (cb1.Checked == true)
            {
              MyChange.FindAndReplace(word, "C444", txtmarka2.Text + " " + txttip2.Text);
              MyChange.FindAndReplace(word, "C555", "Qeydiyyat nişanı – " + txtnomre2.Text);
              MyChange.FindAndReplace(word, "C666", "Qeydiyyat şəhadətnaməsi – " + txtsehadetname2.Text);
              MyChange.FindAndReplace(word, "C6666666", "Zavod - " + txtzavodnomresi2.Text);
              MyChange.FindAndReplace(word, "C777", "Şassi – " + txtsassi2.Text);
              MyChange.FindAndReplace(word, "C888", "Ban – " + txtban2.Text);
              MyChange.FindAndReplace(word, "C999", "Mühərrik – " + txtmuherrik2.Text);
              MyChange.FindAndReplace(word, "C1111", "Rəngi – " + txtrengi2.Text);
              MyChange.FindAndReplace(word, "C2222", "Buraxılış ili - " + txtburaxilis2.Text);
            }
            else
            {
              MyChange.FindAndReplace(word, "C444", "");
              MyChange.FindAndReplace(word, "C555", "");
              MyChange.FindAndReplace(word, "C666", "");
              MyChange.FindAndReplace(word, "C6666666", "");
              MyChange.FindAndReplace(word, "C777", "");
              MyChange.FindAndReplace(word, "C888", "");
              MyChange.FindAndReplace(word, "C999", "");
              MyChange.FindAndReplace(word, "C1111", "");
              MyChange.FindAndReplace(word, "C2222", "");
            }

          MyChange.FindAndReplace(word, "3333", cbmuqavileFormasi.Text);
          MyChange.FindAndReplace(word, "3333", cbmuqavileFormasi.Text);
          MyChange.FindAndReplace(word, "3333", cbmuqavileFormasi.Text);
          MyChange.FindAndReplace(word, "333333333", txtdirektor.Text);
          MyChange.FindAndReplace(word, "333333333", txtdirektor.Text);

            try
            {
                if (comboBox1.Text == "«AGBank» ASC")
                {
                  MyChange.FindAndReplace(word, "111111", "«AGBank» ASC-nin İdarə heyətinin sədri Ə.S Cəlilov (17.08.2017 tarixli etibarnaməyə əsasən) «AGLizinq» QSC-nin Baş Meneceri R.İ. Musayev şəxsində");
                  MyChange.FindAndReplace(word, "111112", comboBox1.Text);
                  MyChange.FindAndReplace(word, "111112", comboBox1.Text);
                  MyChange.FindAndReplace(word, "111112", comboBox1.Text);
                }
                else
                {
                  MyChange.FindAndReplace(word, "111111", "Baş Menecer: Musayev Rəşad İslam oğlu şəxsində");
                  MyChange.FindAndReplace(word, "111112", comboBox1.Text);
                  MyChange.FindAndReplace(word, "111112", comboBox1.Text);
                  MyChange.FindAndReplace(word, "111112", comboBox1.Text);
                }


            }
            catch { }

            doc.Save();

            emeliyyatUcun = emeliyyatUcun + 1;

            if (emeliyyatUcun == 2)
            {
                MyData.insertCommand("baza.accdb", "insert into Emeliyyatlar (a1, a2, a3) Values ('" + dt.Date + "', 'Mülkiyyətə vermə - " + txtnomre.Text + " - " + txtlizinqalan.Text + " - " + txtmarka.Text + "','" + Environment.MachineName + "')");
            }
        }

        public void DYPemr()
        {
            try { Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text); File.Copy("Mulkiyyete verme\\DYP emr.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text + "\\DYP emr - " + txtnomre.Text + ".doc", true); }
            catch { MessageBox.Show("'DYP emr.doc' tapılmadı."); }

            object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text + "\\DYP emr - " + txtnomre.Text + ".doc";

            Word.Application word = new Word.Application();
            Word.Document doc = null;
            object missing = System.Type.Missing;
            object readOnly = false;
            object isVisible = false;
            word.Visible = true;

            doc = word.Documents.Open(ref FileName);
            doc.Activate();


            DateTime dt2 = dttarix2.Value.Date;
            DateTime dt = dttarix.Value.Date;

            string a = MyChange.TarixSozle(dt2);
            string b = MyChange.TarixSozle(dt);

          MyChange.FindAndReplace(word, "0000", dttarix.Text.Substring(0, 2) + " " + b + " " + dttarix.Text.Substring(6, 4) + "- cu il");
          MyChange.FindAndReplace(word, "000", txtlayihe.Text);
          MyChange.FindAndReplace(word, "111", dttarix2.Text.Substring(0, 2) + " " + a + " " + dttarix2.Text.Substring(6, 4) + "- cu il");
          MyChange.FindAndReplace(word, "222", txtlizinqalan.Text + "na");
            if (cb1.Checked == true) { MyChange.FindAndReplace(word, "333", "1 ədəd " + txtmarka.Text + " markalı " + txttip.Text + " avtomobili və 1 ədəd " + txtmarka2.Text + " markalı " + txttip2.Text + " avtomobili"); }
            else MyChange.FindAndReplace(word, "333", "1 ədəd " + txtmarka.Text + " markalı " + txttip.Text + " avtomobili");
          MyChange.FindAndReplace(word, "444", cbmuqavileFormasi.Text);
          MyChange.FindAndReplace(word, "555", comboBox1.Text);

            doc.Save();
        }

        public void DYPemrBANK()
        {
            try { Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text); File.Copy("Mulkiyyete verme\\DYP emr BANK.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text + "\\DYP emr - " + txtnomre.Text + ".doc", true); }
            catch { MessageBox.Show("'DYP emr.doc' tapılmadı."); }

            object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text + "\\DYP emr - " + txtnomre.Text + ".doc";

            Word.Application word = new Word.Application();
            Word.Document doc = null;
            object missing = System.Type.Missing;
            object readOnly = false;
            object isVisible = false;
            word.Visible = true;

            doc = word.Documents.Open(ref FileName);
            doc.Activate();

            DateTime dt2 = dttarix2.Value.Date;
            DateTime dt = dttarix.Value.Date;

            string a = MyChange.TarixSozle(dt2);
            string b = MyChange.TarixSozle(dt);

            MyChange.FindAndReplace(word, "0000", dttarix.Text.Substring(0, 2) + " " + b + " " + dttarix.Text.Substring(6, 4) + "- cu il");
          MyChange.FindAndReplace(word, "000", txtlayihe.Text);
          MyChange.FindAndReplace(word, "111", dttarix2.Text.Substring(0, 2) + " " + a + " " + dttarix2.Text.Substring(6, 4) + "- cu il");
          MyChange.FindAndReplace(word, "222", txtlizinqalan.Text + "na");
            if (cb1.Checked == true) { MyChange.FindAndReplace(word, "333", "1 ədəd " + txtmarka.Text + " markalı " + txttip.Text + " avtomobili və 1 ədəd " + txtmarka2.Text + " markalı " + txttip2.Text + " avtomobili"); }
            MyChange.FindAndReplace(word, "333", "1 ədəd " + txtmarka.Text + " markalı " + txttip.Text + " avtomobili");
          MyChange.FindAndReplace(word, "444", cbmuqavileFormasi.Text);
          MyChange.FindAndReplace(word, "555", comboBox1.Text);

            doc.Save();
        }

        public void DYPmektub()
        {

            try { Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text); File.Copy("Mulkiyyete verme\\DYP mektub.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text + "\\DYP mektub - " + txtnomre.Text + ".doc", true); }
            catch { MessageBox.Show("'DYP mektub.doc' tapılmadı."); }

            object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text + "\\DYP mektub - " + txtnomre.Text + ".doc";


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


            string ban1="", ban2="";

            DateTime dt2 = dttarix2.Value.Date;
            DateTime dt = dttarix.Value.Date;

            string a = MyChange.TarixSozle(dt2);
            string b = MyChange.TarixSozle(dt);

            if (txtsassi.Text != "") ban1 = "Şassi:" + txtsassi.Text; else ban1 = "BAN:" + txtban.Text;

          MyChange.FindAndReplace(word, "0000", dttarix.Text.Substring(0, 2) + " " + b + " " + dttarix.Text.Substring(6, 4) + "- cu il");
          MyChange.FindAndReplace(word, "0000", dttarix.Text.Substring(0, 2) + " " + b + " " + dttarix.Text.Substring(6, 4) + "- cu il");
          MyChange.FindAndReplace(word, "000", txtlayihe.Text);
          MyChange.FindAndReplace(word, "111", dttarix2.Text.Substring(0, 2) + " " + a + " " + dttarix2.Text.Substring(6, 4) + "- cu il");
          MyChange.FindAndReplace(word, "222", txtlizinqalan.Text + "nun");
          MyChange.FindAndReplace(word, "222", txtlizinqalan.Text + "nun");
          MyChange.FindAndReplace(word, "222", txtlizinqalan.Text);

            if (cb1.Checked == true) 
            {
              MyChange.FindAndReplace(word, "333", "1 ədəd " + txtnomre.Text + " dövlət qeydiyyat nişanlı " + txtmarka.Text + " markalı " + txttip.Text + " avtomobili (Tipi:" + txtmodel.Text + ", buraxılış ili:" + txtburaxilis.Text + ", Qeydiyyat Şəhadətnaməsi:" + txtsehadetname.Text + ", mühərrik:" + txtmuherrik.Text + ", " + ban1 + ", Rəngi:" + txtrengi.Text + ")");
              MyChange.FindAndReplace(word, "3333", "və 1 ədəd " + txtnomre2.Text + " dövlət qeydiyyat nişanlı " + txtmarka2.Text + " markalı " + txttip2.Text + " avtomobili (Tipi:" + txtmodel2.Text + ", buraxılış ili:" + txtburaxilis2.Text + ", Qeydiyyat Şəhadətnaməsi:" + txtsehadetname2.Text + ", mühərrik:" + txtmuherrik2.Text + ", " + ban2 + ", Rəngi:" + txtrengi2.Text + ")");
              MyChange.FindAndReplace(word, "444", "1 ədəd " + txtnomre.Text + " dövlət qeydiyyat nişanlı " + txtmarka.Text + " markalı " + txttip.Text + " avtomobili");
              MyChange.FindAndReplace(word, "4444", "və 1 ədəd " + txtnomre2.Text + " dövlət qeydiyyat nişanlı " + txtmarka2.Text + " markalı " + txttip2.Text + " avtomobili");
            }
            else
            {
              MyChange.FindAndReplace(word, "333", "1 ədəd " + txtnomre.Text + " dövlət qeydiyyat nişanlı " + txtmarka.Text + " markalı " + txttip.Text + " avtomobili");
              MyChange.FindAndReplace(word, "444", "1 ədəd " + txtnomre.Text + " dövlət qeydiyyat nişanlı " + txtmarka.Text + " markalı " + txttip.Text + " avtomobili");
              MyChange.FindAndReplace(word, "3333", "");
              MyChange.FindAndReplace(word, "4444", "");
            }

          MyChange.FindAndReplace(word, "2222", cbmuqavileFormasi.Text);

            doc.Save();
        }

        public void DYPmektubBANK()
        {
            try { Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text); File.Copy("Mulkiyyete verme\\DYP mektub BANK.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text + "\\DYP mektub - " + txtnomre.Text + ".doc", true); }
            catch { MessageBox.Show("'DYP mektub.doc' tapılmadı."); }

            object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text + "\\DYP mektub - " + txtnomre.Text + ".doc";


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


            string ban1 = "", ban2 = "";

            DateTime dt2 = dttarix2.Value.Date;
            DateTime dt = dttarix.Value.Date;

            string a = MyChange.TarixSozle(dt2);
            string b = MyChange.TarixSozle(dt);

            MyChange.FindAndReplace(word, "0000", dttarix.Text.Substring(0, 2) + " " + b + " " + dttarix.Text.Substring(6, 4) + "- cu il");
          MyChange.FindAndReplace(word, "0000", dttarix.Text.Substring(0, 2) + " " + b + " " + dttarix.Text.Substring(6, 4) + "- cu il");
          MyChange.FindAndReplace(word, "000", txtlayihe.Text);
          MyChange.FindAndReplace(word, "111", dttarix2.Text.Substring(0, 2) + " " + a + " " + dttarix2.Text.Substring(6, 4) + "- cu il");
          MyChange.FindAndReplace(word, "222", txtlizinqalan.Text + "nun");
          MyChange.FindAndReplace(word, "222", txtlizinqalan.Text + "nun");
          MyChange.FindAndReplace(word, "222", txtlizinqalan.Text);

            if (cb1.Checked == true) 
            {
              MyChange.FindAndReplace(word, "333", "1 ədəd " + txtnomre.Text + " dövlət qeydiyyat nişanlı " + txtmarka.Text + " markalı " + txttip.Text + " avtomobili (Tipi:" + txtmodel.Text + ", buraxılış ili:" + txtburaxilis.Text + ", Qeydiyyat Şəhadətnaməsi:" + txtsehadetname.Text + ", mühərrik:" + txtmuherrik.Text + ", " + ban1 + ", Rəngi:" + txtrengi.Text + ")");
              MyChange.FindAndReplace(word, "3333", "və 1 ədəd " + txtnomre2.Text + " dövlət qeydiyyat nişanlı " + txtmarka2.Text + " markalı " + txttip2.Text + " avtomobili (Tipi:" + txtmodel2.Text + ", buraxılış ili:" + txtburaxilis2.Text + ", Qeydiyyat Şəhadətnaməsi:" + txtsehadetname2.Text + ", mühərrik:" + txtmuherrik2.Text + ", " + ban2 + ", Rəngi:" + txtrengi2.Text + ")");
              MyChange.FindAndReplace(word, "444", "1 ədəd " + txtnomre.Text + " dövlət qeydiyyat nişanlı " + txtmarka.Text + " markalı " + txttip.Text + " avtomobili");
              MyChange.FindAndReplace(word, "4444", "və 1 ədəd " + txtnomre2.Text + " dövlət qeydiyyat nişanlı " + txtmarka2.Text + " markalı " + txttip2.Text + " avtomobili");
            }
            else
            {
              MyChange.FindAndReplace(word, "333", "1 ədəd " + txtnomre.Text + " dövlət qeydiyyat nişanlı " + txtmarka.Text + " markalı " + txttip.Text + " avtomobili");
              MyChange.FindAndReplace(word, "444", "1 ədəd " + txtnomre.Text + " dövlət qeydiyyat nişanlı " + txtmarka.Text + " markalı " + txttip.Text + " avtomobili");
              MyChange.FindAndReplace(word, "3333", "");
              MyChange.FindAndReplace(word, "4444", "");
            }
            /*
          MyChange.FindAndReplace(word, "333", txtnomre.Text);
          MyChange.FindAndReplace(word, "333", txtnomre.Text);
          MyChange.FindAndReplace(word, "444", txtmarka.Text + " markalı " + txttip.Text + " avtomobili");
          MyChange.FindAndReplace(word, "444", txtmarka.Text + " markalı " + txttip.Text + " avtomobili");
          MyChange.FindAndReplace(word, "555", txtsedan.Text);
          MyChange.FindAndReplace(word, "666", txtburaxilis.Text);
          MyChange.FindAndReplace(word, "777", txtsehadetname.Text);
          MyChange.FindAndReplace(word, "888", txtmuherrik.Text);
            try { if (txtsassi.Text != "") this.FindAndReplace(word, "999", "Şassi: " + txtsassi.Text); else this.FindAndReplace(word, "999", "BAN: " + txtban.Text); }
            catch { }
          MyChange.FindAndReplace(word, "1111", txtrengi.Text);*/
          MyChange.FindAndReplace(word, "2222", cbmuqavileFormasi.Text);

            doc.Save();
        }

        public void DYPErize()
        {
            try
            {
                Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text);
                File.Copy("Mulkiyyete verme\\Erize.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text + "\\Erize - " + txtnomre.Text + ".xlsx", true);
            }
            catch { MessageBox.Show("Erize.xlsx tapılmadı."); }


            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text + "\\Erize - " + txtnomre.Text + ".xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

            try
            {
                oSheet.Cells[11, 8] = tel1.Text;
                oSheet.Cells[20, 6] = txtmarka.Text;
                oSheet.Cells[20, 1] = txttip.Text;
                oSheet.Cells[22, 1] = txtmodel.Text;
                oSheet.Cells[22, 6] = "'" + txtzavod.Text;
                oSheet.Cells[24, 1] = txtburaxilis.Text;
                oSheet.Cells[24, 6] = "'" + txtmuherrik.Text;
                oSheet.Cells[26, 1] = "'" + txtban.Text;
                oSheet.Cells[26, 6] = "'" + txtsassi.Text;
                oSheet.Cells[27, 1] = "9. Maksimum kütləsi:  " + txtmaxkutle.Text + " kq";
                oSheet.Cells[28, 6] = txtrengi.Text;
                oSheet.Cells[29, 1] = "11. Yüksüz kütləsi:  " + txtyuksuzkutle.Text + " kq";
                oSheet.Cells[29, 4] = "M.İ.H. " + txtmih.Text + " sm3";
                oSheet.Cells[30, 6] = txtnomre.Text;
                oSheet.Cells[32, 1] = txtsehadetname.Text;
                oSheet.Cells[32, 6] = txttranzit.Text;
                oSheet.Cells[34, 1] = txtlizinqalan.Text;
                if (comboBox1.Text == "«AGLizinq» QSC")
                {
                    oSheet.Cells[5, 1] = "                                             ”AGLizinq” QSC                                               tərəfindən";
                    oSheet.Cells[8, 1] = "Hüquqi şəxsin ünvanı: Landau 16";
                    oSheet.Cells[9, 1] = "Rayon: Yasamal                                                  telefen №-si   012 497 50 17";
                }
                if (comboBox1.Text == "«AGBank» ASC")
                {
                    oSheet.Cells[5, 1] = "                                             ”AGBank” ASC                                               tərəfindən";
                    oSheet.Cells[8, 1] = "Hüquqi şəxsin ünvanı  AZ 1022, Bakı şəhəri Nəsimi rayonu, C.Məmmədquluzadə, ev 102A";
                    oSheet.Cells[9, 1] = "Rayon: Nəsimi                                                   telefen №-si   012 497 50 17";
                }
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

        public void DYPErize2()
        {
            if (txtnomre2.Text != "")
            {
                try
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text);
                    File.Copy("Mulkiyyete verme\\Erize.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text + "\\Erize - " + txtnomre2.Text + ".xlsx", true);
                }
                catch { MessageBox.Show("Erize.xlsx tapılmadı."); }


                //Get a new workbook.
                oXL = new Excel.Application();
                oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text + "\\Erize - " + txtnomre2.Text + ".xlsx"));
                oSheet = (Excel._Worksheet)oWB.Sheets[1];
                oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                oSheet.Activate();
                oSheet.Range["A1"].Select();
                oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

                try
                {
                    oSheet.Cells[11, 8] = tel1.Text;
                    oSheet.Cells[20, 6] = txtmarka2.Text;
                    oSheet.Cells[20, 1] = txttip2.Text;
                    oSheet.Cells[22, 1] = txtmodel2.Text;
                    oSheet.Cells[22, 6] = "'" + txtzavod2.Text;
                    oSheet.Cells[24, 1] = txtburaxilis2.Text;
                    oSheet.Cells[24, 6] = "'" + txtmuherrik2.Text;
                    oSheet.Cells[26, 1] = "'" + txtban2.Text;
                    oSheet.Cells[26, 6] = "'" + txtsassi2.Text;
                    oSheet.Cells[27, 1] = "9. Maksimum kütləsi:  " + txtmaxkutle2.Text + " kq";
                    oSheet.Cells[28, 6] = txtrengi2.Text;
                    oSheet.Cells[29, 1] = "11. Yüksüz kütləsi:  " + txtyuksuzkutle2.Text + " kq";
                    oSheet.Cells[29, 4] = "M.İ.H. " + txtmih2.Text + " sm3";
                    oSheet.Cells[30, 6] = txtnomre2.Text;
                    oSheet.Cells[32, 1] = txtsehadetname2.Text;
                    oSheet.Cells[32, 6] = txttranzit2.Text;
                    oSheet.Cells[34, 1] = txtlizinqalan.Text;
                    if (comboBox1.Text == "«AGLizinq» QSC")
                    {
                        oSheet.Cells[5, 1] = "                                             ”AGLizinq” QSC                                               tərəfindən";
                        oSheet.Cells[8, 1] = "Hüquqi şəxsin ünvanı: Landau 16";
                        oSheet.Cells[9, 1] = "Rayon: Yasamal                                                  telefen №-si   012 497 50 17";
                    }
                    if (comboBox1.Text == "«AGBank» ASC")
                    {
                        oSheet.Cells[5, 1] = "                                             ”AGBank” ASC                                               tərəfindən";
                        oSheet.Cells[8, 1] = "Hüquqi şəxsin ünvanı  AZ 1022, Bakı şəhəri Nəsimi rayonu, C.Məmmədquluzadə, ev 102A";
                        oSheet.Cells[9, 1] = "Rayon: Nəsimi                                                   telefen №-si   012 497 50 17";
                    }
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
        }

        public void DYPTehvilTehvilTeslim()
        {
            try
            {
                Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text);
                File.Copy("Mulkiyyete verme\\Tehvil-Teslim.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text + "\\Tehvil-Teslim - " + txtnomre.Text + ".xlsx", true);
            }
            catch { MessageBox.Show("Tehvil-Teslim.xlsx tapılmadı."); }

            //Get a new workbook.
            oXL2 = new Excel.Application();
            oWB2 = (Excel._Workbook)(oXL2.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text + "\\Tehvil-Teslim - " + txtnomre.Text + ".xlsx"));
            oSheet2 = (Excel._Worksheet)oWB2.Sheets[1];
            oSheet2.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            oSheet2.Activate();
            oSheet2.Range["A1"].Select();
            oSheet2.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;

            try
            {
                oSheet2.Cells[11, 1] = txttip.Text;
                oSheet2.Cells[11, 7] = txtmodel.Text;
                oSheet2.Cells[11, 13] = txtburaxilis.Text;
                oSheet2.Cells[11, 16] = "'" + txtzavod.Text;
                oSheet2.Cells[11, 19] = txtmaxkutle.Text;
                oSheet2.Cells[11, 22] = txtyuksuzkutle.Text;
                oSheet2.Cells[11, 25] = txtsehadetname.Text;

                oSheet2.Cells[18, 1] = txtmarka.Text;
                oSheet2.Cells[18, 7] = "'" + txtban.Text;
                oSheet2.Cells[18, 12] = "'" + txtmuherrik.Text;
                oSheet2.Cells[18, 15] = "'" + txtsassi.Text;
                oSheet2.Cells[18, 20] = txtrengi.Text;
                oSheet2.Cells[18, 23] = txtnomre.Text;
                oSheet2.Cells[18, 26] = txttranzit.Text;
                oSheet2.Cells[49, 5] = txtlizinqalan.Text;

                if (cb1.Checked == true)
                {
                    oSheet2.Cells[12, 1] = txttip2.Text;
                    oSheet2.Cells[12, 7] = txtmodel2.Text;
                    oSheet2.Cells[12, 13] = txtburaxilis2.Text;
                    oSheet2.Cells[12, 16] = "'" + txtzavod2.Text;
                    oSheet2.Cells[12, 19] = txtmaxkutle2.Text;
                    oSheet2.Cells[12, 22] = txtyuksuzkutle2.Text;
                    oSheet2.Cells[12, 25] = txtsehadetname2.Text;

                    oSheet2.Cells[19, 1] = txtmarka2.Text;
                    oSheet2.Cells[19, 7] = "'" + txtban2.Text;
                    oSheet2.Cells[19, 12] = "'" + txtmuherrik2.Text;
                    oSheet2.Cells[19, 15] = "'" + txtsassi2.Text;
                    oSheet2.Cells[19, 20] = txtrengi2.Text;
                    oSheet2.Cells[19, 23] = txtnomre2.Text;
                    oSheet2.Cells[19, 26] = txttranzit2.Text;
                }

                if (comboBox1.Text == "«AGLizinq» QSC")
                {
                    oSheet2.Cells[46, 5] = "“AGLizinq” QSC     Musayev R.İ";
                    oSheet2.Cells[2, 1] = "“AGLizinq” QSC";
                }
                if (comboBox1.Text == "«AGBank» ASC")
                {
                    oSheet2.Cells[46, 5] = "“AGBank” ASC     Musayev R.İ";
                    oSheet2.Cells[2, 1] = "“AGBank” ASC";
                }

            }
            catch { };

            oXL2.Visible = true;
            try
            {
                oXL2.DisplayAlerts = false;
                oWB2.Save();
            }
            catch { }
            // oXL.Application.Quit();
            // oXL.Visible = false;
            // oSheet.PrintOut();
            // oWB.Close(SaveChanges: false);
            // oXL.Workbooks.Close();
        }

        public void DTNMemr()
        {
            try { Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text); File.Copy("Mulkiyyete verme\\DTNM Emr.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text + "\\DTNM Emr - " + txtnomre.Text + ".doc", true); }
            catch { MessageBox.Show("'DTNM Emr.doc' tapılmadı."); }

            object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text + "\\DTNM Emr - " + txtnomre.Text + ".doc";


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


            DateTime dt2 = dttarix2.Value.Date;
            DateTime dt = dttarix.Value.Date;

            string a = MyChange.TarixSozle(dt2);
            string b = MyChange.TarixSozle(dt);

            MyChange.FindAndReplace(word, "0000", dttarix.Text.Substring(0, 2) + " " + b + " " + dttarix.Text.Substring(6, 4) + "- cu il");
          MyChange.FindAndReplace(word, "000", txtlayihe.Text);
          MyChange.FindAndReplace(word, "111", dttarix2.Text.Substring(0, 2) + " " + a + " " + dttarix2.Text.Substring(6, 4) + "- cu il");
          MyChange.FindAndReplace(word, "222", txtlizinqalan.Text + "na");
          MyChange.FindAndReplace(word, "333", cbmuqavileFormasi.Text);
            if (cb1.Checked == true) { MyChange.FindAndReplace(word, "444", "1 ədəd " + txtmarka.Text + " markalı " + txttip.Text + " avtomobili və 1 ədəd " + txtmarka2.Text + " markalı " + txttip2.Text + " avtomobili"); }
            else MyChange.FindAndReplace(word, "444", "1 ədəd " + txtmarka.Text + " markalı " + txttip.Text + " avtomobili");
          MyChange.FindAndReplace(word, "555", comboBox1.Text);

            doc.Save();
        }

        public void DTNMemrBANK()
        {
            try { Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text); File.Copy("Mulkiyyete verme\\DTNM Emr BANK.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text + "\\DTNM Emr - " + txtnomre.Text + ".doc", true); }
            catch { MessageBox.Show("'DTNM Emr.doc' tapılmadı."); }

            object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text + "\\DTNM Emr - " + txtnomre.Text + ".doc";


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

            DateTime dt2 = dttarix2.Value.Date;

            string a = MyChange.TarixSozle(dt2);


            MyChange.FindAndReplace(word, "000", txtlayihe.Text);
          MyChange.FindAndReplace(word, "111", dttarix2.Text.Substring(0, 2) + " " + a + " " + dttarix2.Text.Substring(6, 4) + "- cu il");
          MyChange.FindAndReplace(word, "222", txtlizinqalan.Text + "na");
          MyChange.FindAndReplace(word, "333", cbmuqavileFormasi.Text);
            if (cb1.Checked == true) { MyChange.FindAndReplace(word, "444", "1 ədəd " + txtmarka.Text + " markalı " + txttip.Text + " avtomobili və 1 ədəd " + txtmarka2.Text + " markalı " + txttip2.Text + " avtomobili"); }
            else MyChange.FindAndReplace(word, "444", "1 ədəd " + txtmarka.Text + " markalı " + txttip.Text + " avtomobili");
          MyChange.FindAndReplace(word, "555", comboBox1.Text);

            doc.Save();
        }

        public void DTNMmektub()
        {
            try { Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text); File.Copy("Mulkiyyete verme\\DTNM mektub.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text + "\\DTNM mektub - " + txtnomre.Text + ".doc", true); }
            catch { MessageBox.Show("'DTNM mektub.doc' tapılmadı."); }

            object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text + "\\DTNM mektub - " + txtnomre.Text + ".doc";


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

            DateTime dt2 = dttarix2.Value.Date;
            DateTime dt = dttarix.Value.Date;

            string a = MyChange.TarixSozle(dt2);
            string b = MyChange.TarixSozle(dt);

            MyChange.FindAndReplace(word, "0000", dttarix.Text.Substring(0, 2) + " " + b + " " + dttarix.Text.Substring(6, 4) + "- cu il");
          MyChange.FindAndReplace(word, "0000", dttarix.Text.Substring(0, 2) + " " + b + " " + dttarix.Text.Substring(6, 4) + "- cu il");
          MyChange.FindAndReplace(word, "000", txtlayihe.Text);
          MyChange.FindAndReplace(word, "000", txtlayihe.Text);
          MyChange.FindAndReplace(word, "111", dttarix2.Text.Substring(0, 2) + " " + a + " " + dttarix2.Text.Substring(6, 4) + "- cu il");
          MyChange.FindAndReplace(word, "222", txtlizinqalan.Text + "nun");
          MyChange.FindAndReplace(word, "222", txtlizinqalan.Text + "nun");
          MyChange.FindAndReplace(word, "222", txtlizinqalan.Text);
          MyChange.FindAndReplace(word, "2222", cbmuqavileFormasi.Text);

            if (cb1.Checked == true) 
            {
              MyChange.FindAndReplace(word, "333", "1 ədəd " + txtnomre.Text + " dövlət qeydiyyat nişanlı " + txtmarka.Text + " markalı " + txttip.Text + " (Buraxılış ili:" + txtburaxilis.Text + ", Qeydiyyat Şəhadətnaməsi:" + txtsehadetname.Text + ", Zavod:" + txtzavod.Text + ", Rəngi:" + txtrengi.Text + ")");
              MyChange.FindAndReplace(word, "3333", "və 1 ədəd " + txtnomre2.Text + " dövlət qeydiyyat nişanlı " + txtmarka2.Text + " markalı " + txttip2.Text + " (Buraxılış ili:" + txtburaxilis2.Text + ", Qeydiyyat Şəhadətnaməsi:" + txtsehadetname2.Text + ", Zavod:" + txtzavod2.Text + ", Rəngi:" + txtrengi2.Text + ")");
              MyChange.FindAndReplace(word, "444", "1 ədəd " + txtnomre.Text + " dövlət qeydiyyat nişanlı " + txtmarka.Text + " markalı " + txttip.Text);
              MyChange.FindAndReplace(word, "4444", "və 1 ədəd " + txtnomre2.Text + " dövlət qeydiyyat nişanlı " + txtmarka2.Text + " markalı " + txttip2.Text);
            }
            else
            {
              MyChange.FindAndReplace(word, "333", "1 ədəd " + txtnomre.Text + " dövlət qeydiyyat nişanlı " + txtmarka.Text + " markalı " + txttip.Text);
              MyChange.FindAndReplace(word, "444", "1 ədəd " + txtnomre.Text + " dövlət qeydiyyat nişanlı " + txtmarka.Text + " markalı " + txttip.Text);
              MyChange.FindAndReplace(word, "3333", "");
              MyChange.FindAndReplace(word, "4444", "");
            }
            
            doc.Save();
        }

        public void DTNMmektubBANK()
        {
            try { Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text); File.Copy("Mulkiyyete verme\\DTNM mektub BANK.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text + "\\DTNM mektub - " + txtnomre.Text + ".doc", true); }
            catch { MessageBox.Show("'DTNM mektub.doc' tapılmadı."); }

            object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text + "\\DTNM mektub - " + txtnomre.Text + ".doc";


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

            DateTime dt2 = dttarix2.Value.Date;
            DateTime dt = dttarix.Value.Date;

            string a = MyChange.TarixSozle(dt2);
            string b = MyChange.TarixSozle(dt);

            MyChange.FindAndReplace(word, "0000", dttarix.Text.Substring(0, 2) + " " + b + " " + dttarix.Text.Substring(6, 4) + "- cu il");
          MyChange.FindAndReplace(word, "000", txtlayihe.Text);
          MyChange.FindAndReplace(word, "000", txtlayihe.Text);
          MyChange.FindAndReplace(word, "111", dttarix2.Text.Substring(0, 2) + " " + a + " " + dttarix2.Text.Substring(6, 4) + "- cu il");
          MyChange.FindAndReplace(word, "222", txtlizinqalan.Text + "nun");
          MyChange.FindAndReplace(word, "222", txtlizinqalan.Text + "nun");
          MyChange.FindAndReplace(word, "222", txtlizinqalan.Text);
          MyChange.FindAndReplace(word, "2222", cbmuqavileFormasi.Text);

            if (cb1.Checked == true) 
            {
              MyChange.FindAndReplace(word, "333", "1 ədəd " + txtnomre.Text + " dövlət qeydiyyat nişanlı " + txtmarka.Text + " markalı " + txttip.Text + " (Buraxılış ili:" + txtburaxilis.Text + ", Qeydiyyat Şəhadətnaməsi:" + txtsehadetname.Text + ", Zavod:" + txtzavod.Text + ", Rəngi:" + txtrengi.Text + ")");
              MyChange.FindAndReplace(word, "3333", "və 1 ədəd " + txtnomre2.Text + " dövlət qeydiyyat nişanlı " + txtmarka2.Text + " markalı " + txttip2.Text + " (Buraxılış ili:" + txtburaxilis2.Text + ", Qeydiyyat Şəhadətnaməsi:" + txtsehadetname2.Text + ", Zavod:" + txtzavod2.Text + ", Rəngi:" + txtrengi2.Text + ")");
              MyChange.FindAndReplace(word, "444", "1 ədəd " + txtnomre.Text + " dövlət qeydiyyat nişanlı " + txtmarka.Text + " markalı " + txttip.Text);
              MyChange.FindAndReplace(word, "4444", "və 1 ədəd " + txtnomre2.Text + " dövlət qeydiyyat nişanlı " + txtmarka2.Text + " markalı " + txttip2.Text);
            }
            else
            {
              MyChange.FindAndReplace(word, "333", "1 ədəd " + txtnomre.Text + " dövlət qeydiyyat nişanlı " + txtmarka.Text + " markalı " + txttip.Text);
              MyChange.FindAndReplace(word, "444", "1 ədəd " + txtnomre.Text + " dövlət qeydiyyat nişanlı " + txtmarka.Text + " markalı " + txttip.Text);
              MyChange.FindAndReplace(word, "3333", "");
              MyChange.FindAndReplace(word, "4444", "");
            }


            doc.Save();
        }

        public void DTNMerize()
        {
            try { Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text); File.Copy("Mulkiyyete verme\\DTNM Erize.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text + "\\DTNM Erize - " + txtnomre.Text + ".doc", true); }
            catch { MessageBox.Show("'DTNM Erize.doc' tapılmadı."); }

            object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text + "\\DTNM Erize - " + txtnomre.Text + ".doc";


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

          MyChange.FindAndReplace(word, "0000", dttarix.Text + "- cı il");
          MyChange.FindAndReplace(word, "0000", dttarix.Text + "- cı il");
          MyChange.FindAndReplace(word, "0000", dttarix.Text + "- cı il");
          MyChange.FindAndReplace(word, "333", txtnomre.Text);
          MyChange.FindAndReplace(word, "444", txtmarka.Text);
          MyChange.FindAndReplace(word, "444", txtmarka.Text);
          MyChange.FindAndReplace(word, "555", comboBox1.Text);
          MyChange.FindAndReplace(word, "555", comboBox1.Text);
          MyChange.FindAndReplace(word, "555", comboBox1.Text);
          MyChange.FindAndReplace(word, "555", comboBox1.Text);
          MyChange.FindAndReplace(word, "3331", txtburaxilis.Text);
          MyChange.FindAndReplace(word, "3332", txtsehadetname.Text);
          MyChange.FindAndReplace(word, "3333", txtzavodnomresi.Text);
          MyChange.FindAndReplace(word, "3334", txtrengi.Text);
          MyChange.FindAndReplace(word, "9999", txtsassi.Text);
          MyChange.FindAndReplace(word, "8888", txtmuherrik.Text);
          MyChange.FindAndReplace(word, "222", txtlizinqalan.Text);
          MyChange.FindAndReplace(word, "111114", txtmodel.Text + " " + txttip.Text); //texnikanin novu
          MyChange.FindAndReplace(word, "111115", txtmih.Text); //muherrikin gucu
          MyChange.FindAndReplace(word, "111116", txtzavod.Text); //istehsalci olke

            if (tel1.Text == tel2.Text || tel2.Text == "       -  -") tel2.Text = "";
          MyChange.FindAndReplace(word, "111", tel1.Text + "; " + tel2.Text);

            if (comboBox1.Text == "«AGBank» ASC")
            {
              MyChange.FindAndReplace(word, "111111", "AZ1022 Bakı şəh., Nəsimi r-nu, C.Məmmədquluzadə, ev 102A.");
              MyChange.FindAndReplace(word, "111112", "AZ1022 Bakı şəh., Nəsimi r-nu, C.Məmmədquluzadə, ev 102A.");
              MyChange.FindAndReplace(word, "111113", "9900019651");
            }
            else
            {
              MyChange.FindAndReplace(word, "111111", "AZ1073 Bakı şəhəri, Yasamal r-nu, Landau 16");
              MyChange.FindAndReplace(word, "111112", "AZ 1022, Bakı şəhəri Nəsimi rayonu, C.Məmmədquluzadə, ev 102A");
              MyChange.FindAndReplace(word, "111113", "1300616961");
            }


            doc.Save();
        }

        public void DTNMerize2()
        {
            if (txtnomre2.Text != "")
            {
                try { Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text); File.Copy("Mulkiyyete verme\\DTNM Erize.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text + "\\DTNM Erize - " + txtnomre2.Text + ".doc", true); }
                catch { MessageBox.Show("'DTNM Erize.doc' tapılmadı."); }

                object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text + "\\DTNM Erize - " + txtnomre2.Text + ".doc";


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

              MyChange.FindAndReplace(word, "0000", dttarix.Text + "- cı il");
              MyChange.FindAndReplace(word, "0000", dttarix.Text + "- cı il");
              MyChange.FindAndReplace(word, "0000", dttarix.Text + "- cı il");
              MyChange.FindAndReplace(word, "333", txtnomre2.Text);
              MyChange.FindAndReplace(word, "444", txtmarka2.Text);
              MyChange.FindAndReplace(word, "444", txtmarka2.Text);
              MyChange.FindAndReplace(word, "555", comboBox1.Text);
              MyChange.FindAndReplace(word, "555", comboBox1.Text);
              MyChange.FindAndReplace(word, "555", comboBox1.Text);
              MyChange.FindAndReplace(word, "555", comboBox1.Text);
              MyChange.FindAndReplace(word, "3331", txtburaxilis2.Text);
              MyChange.FindAndReplace(word, "3332", txtsehadetname2.Text);
              MyChange.FindAndReplace(word, "3333", txtzavodnomresi2.Text);
              MyChange.FindAndReplace(word, "3334", txtrengi2.Text);
              MyChange.FindAndReplace(word, "9999", txtsassi2.Text);
              MyChange.FindAndReplace(word, "8888", txtmuherrik2.Text);
              MyChange.FindAndReplace(word, "222", txtlizinqalan.Text);
              MyChange.FindAndReplace(word, "111114", txtmodel2.Text + " " + txttip.Text); //texnikanin novu
              MyChange.FindAndReplace(word, "111115", txtmih2.Text); //muherrikin gucu
              MyChange.FindAndReplace(word, "111116", txtzavod2.Text); //istehsalci olke

                if (tel1.Text == tel2.Text || tel2.Text == "       -  -") tel2.Text = "";
              MyChange.FindAndReplace(word, "111", tel1.Text + "; " + tel2.Text);

                if (comboBox1.Text == "«AGBank» ASC")
                {
                  MyChange.FindAndReplace(word, "111111", "AZ1022 Bakı şəh., Nəsimi r-nu, C.Məmmədquluzadə, ev 102A.");
                  MyChange.FindAndReplace(word, "111112", "AZ1022 Bakı şəh., Nəsimi r-nu, C.Məmmədquluzadə, ev 102A.");
                }
                else
                {
                  MyChange.FindAndReplace(word, "111111", "AZ 1073 Bakı şəhəri, Yasamal r-nu, Landau 16");
                  MyChange.FindAndReplace(word, "111112", "AZ 1022, Bakı şəhəri Nəsimi rayonu, C.Məmmədquluzadə, ev 102A");
                }


                doc.Save();
            }
        }

        public void DTNMTexnikiBaxis()
        {
            try { Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text); File.Copy("Mulkiyyete verme\\DTNM TexnikiBaxis.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text + "\\DTNM TexnikiBaxis - " + txtnomre.Text + ".doc", true); }
            catch { MessageBox.Show("'DTNM Texniki Baxis.doc' tapılmadı."); }

            object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\Mulkiyyete verme " + txtnomre.Text + "\\DTNM TexnikiBaxis - " + txtnomre.Text + ".doc";


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

            DateTime dt = dttarix.Value.Date;
            string b = MyChange.TarixSozle(dt);

            MyChange.FindAndReplace(word, "0000", dttarix.Text.Substring(0, 2) + " " + b + " " + dttarix.Text.Substring(6, 4) + "- cı il");
          MyChange.FindAndReplace(word, "222", txttip.Text);
          MyChange.FindAndReplace(word, "333", txtnomre.Text);
          MyChange.FindAndReplace(word, "444", txtmarka.Text);
          MyChange.FindAndReplace(word, "555", comboBox1.Text);
          MyChange.FindAndReplace(word, "5555", "Lizinq alan " + txtlizinqalan.Text);//oglunun istiraki ile
          MyChange.FindAndReplace(word, "3331", txtburaxilis.Text);
          MyChange.FindAndReplace(word, "3333", txtzavodnomresi.Text);
          MyChange.FindAndReplace(word, "8888", txtmuherrik.Text);

            if (cb1.Checked == true)
            {
              MyChange.FindAndReplace(word, "C222", txttip2.Text);
              MyChange.FindAndReplace(word, "C333", txtnomre2.Text);
              MyChange.FindAndReplace(word, "C444", txtmarka2.Text);
              MyChange.FindAndReplace(word, "C3331", txtburaxilis2.Text);
              MyChange.FindAndReplace(word, "C3333", txtzavodnomresi2.Text);
              MyChange.FindAndReplace(word, "C8888", txtmuherrik2.Text);
            }
            else
            {
              MyChange.FindAndReplace(word, "C222", "");
              MyChange.FindAndReplace(word, "C333", "");
              MyChange.FindAndReplace(word, "C444","");
              MyChange.FindAndReplace(word, "C3331", "");
              MyChange.FindAndReplace(word, "C3333", "");
              MyChange.FindAndReplace(word, "C8888", "");
            }

            
            doc.Save();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            AutoDuzelis();
            
            if (cbhuquqisexs.Text == "Fiziki şəxs") { WordDoc(); }
            else if (cbhuquqisexs.Text == "Hüquqi şəxs") { WordDocHuquqiSexs(); } 
           
        }

        private void txtnomre_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    MyData.selectCommand("baza.accdb", "Select * from etibarnameneqliyyat where c1 Like " + "'%" + txtnomre.Text + "%'");
                    MyData.dtmain = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmain);

                    txtlayihe.Text = MyData.dtmain.Rows[0]["c4"].ToString(); txtlayihe.BackColor = Color.LimeGreen;
                    txtlizinqalan.Text = MyData.dtmain.Rows[0]["c3"].ToString(); txtlizinqalan.BackColor = Color.LimeGreen;
                    txtmarka.Text = MyData.dtmain.Rows[0]["c2"].ToString(); txtmarka.BackColor = Color.LimeGreen;
                    txtnomre.Text = MyData.dtmain.Rows[0]["c1"].ToString(); txtnomre.BackColor = Color.LimeGreen;
                    txtsehadetname.Text = MyData.dtmain.Rows[0]["c12"].ToString(); txtsehadetname.BackColor = Color.LimeGreen;
                    txtsassi.Text = MyData.dtmain.Rows[0]["c10"].ToString(); txtsassi.BackColor = Color.LimeGreen;
                    txtban.Text = MyData.dtmain.Rows[0]["c8"].ToString(); txtban.BackColor = Color.LimeGreen;
                    txtmuherrik.Text = MyData.dtmain.Rows[0]["c9"].ToString(); txtmuherrik.BackColor = Color.LimeGreen;
                    txtrengi.Text = MyData.dtmain.Rows[0]["c5"].ToString(); txtrengi.BackColor = Color.LimeGreen;
                    txtburaxilis.Text = MyData.dtmain.Rows[0]["c6"].ToString(); txtburaxilis.BackColor = Color.LimeGreen;
                    txtzavodnomresi.Text = MyData.dtmain.Rows[0]["c11"].ToString(); txtzavodnomresi.BackColor = Color.LimeGreen;

                    MyData.selectCommand("baza.accdb", "Select * from Telefon where c1 Like '%" + txtlizinqalan.Text + "%'");
                    MyData.dtmain = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmain);

                    tel1.Text = ""; tel2.Text = "";
                    try { tel1.Text = MyData.dtmain.Rows[0]["c2"].ToString(); }
                    catch { };
                    try { tel2.Text = MyData.dtmain.Rows[0]["c3"].ToString(); }
                    catch { };

                    MMXrefresh();
                }
                catch { }
            }
        }

        private void txtlayihe_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtmarka2.Text = "";
                txtnomre2.Text = "";
                txtsehadetname2.Text = "";
                txtsassi2.Text = "";
                txtban2.Text = "";
                txtmuherrik2.Text = "";
                txtrengi2.Text = "";
                txtburaxilis2.Text = "";
                txtzavodnomresi2.Text = "";

                try
                {
                    MyData.selectCommand("baza.accdb", "Select * from etibarnameneqliyyat where c4 Like " + "'%" + txtlayihe.Text + "%'");
                    MyData.dtmain = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmain);

                    txtlayihe.Text = MyData.dtmain.Rows[0]["c4"].ToString();
                    txtlizinqalan.Text = MyData.dtmain.Rows[0]["c3"].ToString();
                    txtmarka.Text = MyData.dtmain.Rows[0]["c2"].ToString();
                    txtnomre.Text = MyData.dtmain.Rows[0]["c1"].ToString();
                    txtsehadetname.Text = MyData.dtmain.Rows[0]["c12"].ToString();
                    txtsassi.Text = MyData.dtmain.Rows[0]["c10"].ToString();
                    txtban.Text = MyData.dtmain.Rows[0]["c8"].ToString();
                    txtmuherrik.Text = MyData.dtmain.Rows[0]["c9"].ToString();
                    txtrengi.Text = MyData.dtmain.Rows[0]["c5"].ToString();
                    txtburaxilis.Text = MyData.dtmain.Rows[0]["c6"].ToString();
                    txtzavodnomresi.Text = MyData.dtmain.Rows[0]["c11"].ToString();

                    try
                    {
                        txtmarka2.Text = MyData.dtmain.Rows[1]["c2"].ToString();
                        txtnomre2.Text = MyData.dtmain.Rows[1]["c1"].ToString();
                        txtsehadetname2.Text = MyData.dtmain.Rows[1]["c12"].ToString();
                        txtsassi2.Text = MyData.dtmain.Rows[1]["c10"].ToString();
                        txtban2.Text = MyData.dtmain.Rows[1]["c8"].ToString();
                        txtmuherrik2.Text = MyData.dtmain.Rows[1]["c9"].ToString();
                        txtrengi2.Text = MyData.dtmain.Rows[1]["c5"].ToString();
                        txtburaxilis2.Text = MyData.dtmain.Rows[1]["c6"].ToString();
                        txtzavodnomresi2.Text = MyData.dtmain.Rows[1]["c11"].ToString();
                    }
                    catch { }
                    MyData.selectCommand("baza.accdb", "Select * from Telefon where c1 Like '%" + txtlizinqalan.Text + "%'");
                    MyData.dtmain = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmain);

                    tel1.Text = ""; tel2.Text = "";
                    try { tel1.Text = MyData.dtmain.Rows[0]["c2"].ToString(); }
                    catch { };
                    try { tel2.Text = MyData.dtmain.Rows[0]["c3"].ToString(); }
                    catch { };
                }
                catch { }
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            DateTime dt = dttarix.Value.Date;
            DateTime dt2 = dttarix2.Value.Date;

            AutoDuzelis();

            if (dt == dt2)
            {
                if (!MyCheck.davamYesNo("Müqavilənin tarixinin düzgünlüyündən əminsinizmi?")) return;
            }

            if (comboBox2.Text == "DÖVLƏTTEXNƏZARƏT MÜFƏTTİŞLİYİ")
            {
                if (comboBox1.Text == "«AGBank» ASC")
                {
                    DTNMemrBANK();
                    DTNMmektubBANK();
                }
                else
                {
                    DTNMemr();
                    DTNMmektub();
                }
                
                    DTNMerize();
                    if (cb1.Checked == true) DTNMerize2();
                    DTNMTexnikiBaxis();
            }
            else
            {
                if (comboBox1.Text == "«AGBank» ASC")
                {
                    DYPemrBANK();
                    DYPmektubBANK();
                }
                else
                {
                    DYPemr();
                    DYPmektub();
                }

                DYPErize();
                if (cb1.Checked == true) DYPErize2();
                DYPTehvilTehvilTeslim();
            }

            emeliyyatUcun = emeliyyatUcun + 1;

            if (emeliyyatUcun == 2)
            {
                MyData.insertCommand("baza.accdb", "insert into Emeliyyatlar (a1, a2, a3) Values ('" + dt.Date + "', 'Mülkiyyətə vermə - " + txtnomre.Text + " - " + txtlizinqalan.Text + " - " + txtmarka.Text + "','" + Environment.MachineName + "')");
            }
        }

        private void txtlizinqalan_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    MyData.selectCommand("baza.accdb", "Select * from etibarnameneqliyyat where c3 Like " + "'%" + txtlizinqalan.Text + "%'");
                    MyData.dtmain = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmain);
                    txtlizinqalan.Text = MyData.dtmain.Rows[0]["c3"].ToString();
                }
                catch { }
            }
        }

        private void comboBox2_TextChanged(object sender, EventArgs e)
        {
            if (comboBox2.Text == "DÖVLƏTTEXNƏZARƏT MÜFƏTTİŞLİYİ") { txttip.Text = "Ekskavator"; txtmodel.Text = ""; }
            if (comboBox2.Text == "DAXİLİ İŞLƏR NAZİRLİYİ") { txttip.Text = "minik"; txtmodel.Text = "sedan"; }
        }

        private void cbhuquqisexs_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbhuquqisexs.Text == "Fiziki şəxs") { label18.Visible = false; txtdirektor.Visible = false; }
            if (cbhuquqisexs.Text == "Hüquqi şəxs") { label18.Visible = true; txtdirektor.Visible = true; }
        }

        private void txtnomre2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    MyData.selectCommand("baza.accdb", "Select * from etibarnameneqliyyat where c1 Like " + "'%" + txtnomre2.Text + "%'");
                    MyData.dtmain = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmain);

                    txtmarka2.Text = MyData.dtmain.Rows[0]["c2"].ToString(); txtmarka2.BackColor = Color.LimeGreen;
                    txtnomre2.Text = MyData.dtmain.Rows[0]["c1"].ToString(); txtnomre2.BackColor = Color.LimeGreen;
                    txtsehadetname2.Text = MyData.dtmain.Rows[0]["c12"].ToString(); txtsehadetname2.BackColor = Color.LimeGreen;
                    txtsassi2.Text = MyData.dtmain.Rows[0]["c10"].ToString(); txtsassi2.BackColor = Color.LimeGreen;
                    txtban2.Text = MyData.dtmain.Rows[0]["c8"].ToString(); txtban2.BackColor = Color.LimeGreen;
                    txtmuherrik2.Text = MyData.dtmain.Rows[0]["c9"].ToString(); txtmuherrik2.BackColor = Color.LimeGreen;
                    txtrengi2.Text = MyData.dtmain.Rows[0]["c5"].ToString(); txtrengi2.BackColor = Color.LimeGreen;
                    txtburaxilis2.Text = MyData.dtmain.Rows[0]["c6"].ToString(); txtburaxilis2.BackColor = Color.LimeGreen;
                    txtzavodnomresi2.Text = MyData.dtmain.Rows[0]["c11"].ToString(); txtzavodnomresi2.BackColor = Color.LimeGreen;
                }
                catch { }
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (cb1.Checked == true)
            {
                txttip2.Visible = true;
                txtmodel2.Visible = true;
                txtmaxkutle2.Visible = true;
                txtyuksuzkutle2.Visible = true;
                txtmih2.Visible = true;
                txtzavod2.Visible = true;
                txttranzit2.Visible = true;

                txtmarka.Width = 210;
                txtnomre.Width = 210;
                txtsehadetname.Width = 210;
                txtsassi.Width = 210;
                txtban.Width = 210;
                txtmuherrik.Width = 210;
                txtrengi.Width = 210;
                txtburaxilis.Width = 210;
                txtzavodnomresi.Width = 210;

                txtmaxkutle.Width = 142;
                txtyuksuzkutle.Width = 142;
                txtmih.Width = 142;
                txtzavod.Width = 142;
                txttranzit.Width = 142;

                txtmarka2.Visible = true;
                txtnomre2.Visible = true;
                txtsehadetname2.Visible = true;
                txtsassi2.Visible = true;
                txtban2.Visible = true;
                txtmuherrik2.Visible = true;
                txtrengi2.Visible = true;
                txtburaxilis2.Visible = true;
                txtzavodnomresi2.Visible = true;
            }
            else
            {
                txttip2.Visible = false;
                txtmodel2.Visible = false;
                txtmaxkutle2.Visible = false;
                txtyuksuzkutle2.Visible = false;
                txtmih2.Visible = false;
                txtzavod2.Visible = false;
                txttranzit2.Visible = false;

                txtmarka.Width = 426;
                txtnomre.Width = 426;
                txtsehadetname.Width = 426;
                txtsassi.Width = 426;
                txtban.Width = 426;
                txtmuherrik.Width = 426;
                txtrengi.Width = 426;
                txtburaxilis.Width = 426;
                txtzavodnomresi.Width = 426;

                txtmaxkutle.Width = 288;
                txtyuksuzkutle.Width = 288;
                txtmih.Width = 288;
                txtzavod.Width = 288;
                txttranzit.Width = 288;

                txtmarka2.Visible = false;
                txtnomre2.Visible = false;
                txtsehadetname2.Visible = false;
                txtsassi2.Visible = false;
                txtban2.Visible = false;
                txtmuherrik2.Visible = false;
                txtrengi2.Visible = false;
                txtburaxilis2.Visible = false;
                txtzavodnomresi2.Visible = false;
            }
        }

        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {
            DateTime dt = dttarix.Value.Date;
            try
            {
                MyData.selectCommand("baza.accdb", "UPDATE MMX SET "
                                                                                     + "a8 ='Bəli',"
                                                                                     + "a9 ='Bəli " + dt.Date+ "'"
                                                                                     + " WHERE NOT a9 Like '%Bəli%' and a2 Like '%" + txtnomre.Text + "%'");
                MessageBox.Show("Protokollar təhvil verildi.");
                MMXrefresh();
            }
            catch { MessageBox.Show("Əməliyyat baş tutmadı."); }
            
        }

        private void btMMX2_Click(object sender, EventArgs e)
        {
            MMX mmx = new MMX();
            mmx.textBox1.Text = txtnomre.Text;
            mmx.Show();
            
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            AutoDuzelis();
        }

        private void Mulkiyyete_Verme_Load(object sender, EventArgs e)
        {

        }
    }
}
