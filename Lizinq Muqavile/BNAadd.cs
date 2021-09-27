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
using System.Globalization;
using System.Net.Mail;
using System.Net;
using System.IO;
using System.Net.Sockets;
using System.Web;
using System.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;
using Nsoft;

namespace Lizinq_Muqavile
{
    public partial class BNAadd : Form
    {

        public string Metn;

        public BNAadd()
        {
            InitializeComponent();
        }

        private void SendMail()
        {
            try
            {
                Outlook.Application oApp = new Outlook.Application();
                Outlook._MailItem oMailItem = (Outlook._MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
                oMailItem.To = cbMailTo.Text;
                oMailItem.CC = "info@bna.az";
                oMailItem.Subject = "Sorğu " + txtSorgu.Text + " / " + t2.Text;
                oMailItem.Body = Metn;
                oMailItem.Display(true); // Outlook send file sehifesin acmaq ucun

                Clipboard.SetText(Metn); // metni clipboarda kopyalamaq ucun
                
            }
            catch { }
        }

        private void t2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    MyData.selectCommand("baza.accdb", "Select * from etibarnameneqliyyat where c1 Like '%" + t2.Text + "%'");
                    MyData.dtmainEtibarbame = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmainEtibarbame);

                    if (MyData.dtmainEtibarbame.Rows[0]["c13"].ToString() == "AGB") comboBox1.Text = "“AGBank” ASC";
                    if (MyData.dtmainEtibarbame.Rows[0]["c13"].ToString() == "AGL") comboBox1.Text = "“AGLizinq” QSC";

                    t2.Text = MyData.dtmainEtibarbame.Rows[0]["c1"].ToString();
                    t3.Text = MyData.dtmainEtibarbame.Rows[0]["c2"].ToString();
                    txtlizinqalan.Text = MyData.dtmainEtibarbame.Rows[0]["c3"].ToString();
                    txtlayihe.Text = MyData.dtmainEtibarbame.Rows[0]["c4"].ToString();

                }
                catch { }


                MyData.selectCommand("baza.accdb", "SELECT * FROM etibarnamearxiv WHERE a16 LIKE '%" + t2.Text + "%'"); //Verilmis  etibarnameler
                MyData.dtmainEtibarnameArxiv = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainEtibarnameArxiv);
                dataGridView1.DataSource = MyData.dtmainEtibarnameArxiv;

                if (dataGridView1.Rows.Count > 0)
                {
                    dataGridView1.Rows[dataGridView1.Rows.Count - 1].Selected = true;
                    checkBox2.Checked = true;
                }
                else checkBox2.Checked = false;

                /////////////////////Lizinq alanin melumatlari dolsun deye
                try
                {
                    lbSur.Text = "Sür. vəsiqəsi"; lbSexsiyyet.Text = "Ş/V"; lbFiziki.Text = "Fiziki şəxs"; lbSur.Left = 9; lbSexsiyyet.Left = 72;

                    MyData.selectCommand("baza.accdb", "Select * from etibarnamesurucu where a1 Like '%" + txtlizinqalan.Text + "%'");
                    MyData.dtmainSurucu = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmainSurucu);

                    if (MyData.dtmainSurucu.Rows.Count > 0 & txtlizinqalan.Text != "")
                    {
                        try { txtlizinqalan.Text = MyData.dtmainSurucu.Rows[0]["a1"].ToString(); }
                        catch { /*t7.Text = "";*/ }
                        try { tel1.Text = MyData.dtmainSurucu.Rows[0]["a8"].ToString(); }
                        catch { /*tel1.Text = "";*/ }
                        try { txtunvan1.Text = MyData.dtmainSurucu.Rows[0]["a4"].ToString(); }
                        catch { /*txtunvan.Text = ""; */}
                        try { txtSurvesiqe1.Text = MyData.dtmainSurucu.Rows[0]["a3"].ToString(); }
                        catch { /*txtunvan.Text = ""; */}
                        try { txtsexsiyyet1.Text = MyData.dtmainSurucu.Rows[0]["a2"].ToString(); }
                        catch { /*txtunvan.Text = ""; */}
                    }
                    else
                    {
                        lbSur.Text = "VÖEN"; lbSexsiyyet.Text = "Direktor"; lbFiziki.Text = "Hüquqi şəxs"; lbSur.Left = 53; lbSexsiyyet.Left = 42;

                        txtunvan1.Text = ""; tel1.Text = ""; txtSurvesiqe1.Text = ""; txtsexsiyyet1.Text = "";

                        MyData.selectCommand("baza.accdb", "Select * from muqavilerekvizit where [Lizinq alan] Like '%" + txtlizinqalan.Text + "%'");
                        MyData.dtmainSurucu = new DataTable();
                        MyData.oledbadapter1.Fill(MyData.dtmainMuqavileRekvizit);

                        if (MyData.dtmainMuqavileRekvizit.Rows.Count > 0 & txtlizinqalan.Text != "")
                        {
                            try { txtlizinqalan.Text = MyData.dtmainMuqavileRekvizit.Rows[0]["Lizinq alan"].ToString(); }
                            catch { /*t7.Text = "";*/ }
                            try { tel1.Text = MyData.dtmainMuqavileRekvizit.Rows[0]["Əlaqə nömrəsi 1"].ToString(); }
                            catch { /*tel1.Text = "";*/ }
                            try { txtunvan1.Text = MyData.dtmainMuqavileRekvizit.Rows[0]["Vöen qeydiyyat ünvanı"].ToString(); }
                            catch { /*txtunvan.Text = ""; */}
                            try { txtSurvesiqe1.Text = MyData.dtmainMuqavileRekvizit.Rows[0]["Vöen"].ToString(); }
                            catch { /*txtunvan.Text = ""; */}
                            try { txtsexsiyyet1.Text = MyData.dtmainMuqavileRekvizit.Rows[0]["Direktor"].ToString(); }
                            catch { /*txtunvan.Text = ""; */}
                        }
                        else { txtunvan1.Text = ""; tel1.Text = ""; txtSurvesiqe1.Text = ""; txtsexsiyyet1.Text = ""; }
                    }
                    MyData.selectCommand("baza.accdb", "Select * from etibarnamesurucu where a5 Like '%" + txtlayihe.Text + "%'");
                    MyData.dtmainSurucu = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmainSurucu);
                    dataGridView2.DataSource = MyData.dtmainSurucu;

                }
                catch { }
            }
        }

        private void txtlizinqalan_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    lbSur.Text = "Sür. vəsiqəsi"; lbSexsiyyet.Text = "Ş/V"; lbFiziki.Text = "Fiziki şəxs"; lbSur.Left = 9; lbSexsiyyet.Left = 72;

                    MyData.selectCommand("baza.accdb", "Select * from etibarnamesurucu where a1 Like '%" + txtlizinqalan.Text + "%'");
                    MyData.dtmainSurucu = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmainSurucu);

                    if (MyData.dtmainSurucu.Rows.Count > 0 & txtlizinqalan.Text != "")
                    {
                            try { txtlizinqalan.Text = MyData.dtmainSurucu.Rows[0]["a1"].ToString(); }
                            catch { /*t7.Text = "";*/ }
                            try { tel1.Text = MyData.dtmainSurucu.Rows[0]["a8"].ToString(); }
                            catch { /*tel1.Text = "";*/ }
                            try { txtunvan1.Text = MyData.dtmainSurucu.Rows[0]["a4"].ToString(); }
                            catch { /*txtunvan.Text = ""; */}
                            try { txtSurvesiqe1.Text = MyData.dtmainSurucu.Rows[0]["a3"].ToString(); }
                            catch { /*txtunvan.Text = ""; */}
                            try { txtsexsiyyet1.Text = MyData.dtmainSurucu.Rows[0]["a2"].ToString(); }
                            catch { /*txtunvan.Text = ""; */}
                    }
                    else
                    {
                        lbSur.Text = "VÖEN"; lbSexsiyyet.Text = "Direktor"; lbFiziki.Text = "Hüquqi şəxs"; lbSur.Left = 53; lbSexsiyyet.Left = 42;

                        txtunvan1.Text = ""; tel1.Text = ""; txtSurvesiqe1.Text = ""; txtsexsiyyet1.Text = "";

                        MyData.selectCommand("baza.accdb", "Select * from muqavilerekvizit where [Lizinq alan] Like '%" + txtlizinqalan.Text + "%'");
                        MyData.dtmainMuqavileRekvizit = new DataTable();
                        MyData.oledbadapter1.Fill(MyData.dtmainMuqavileRekvizit);

                        if (MyData.dtmainMuqavileRekvizit.Rows.Count > 0 & txtlizinqalan.Text != "")
                        {
                            try { txtlizinqalan.Text = MyData.dtmainMuqavileRekvizit.Rows[0]["Lizinq alan"].ToString(); }
                            catch { /*t7.Text = "";*/ }
                            try { tel1.Text = MyData.dtmainMuqavileRekvizit.Rows[0]["Əlaqə nömrəsi 1"].ToString(); }
                            catch { /*tel1.Text = "";*/ }
                            try { txtunvan1.Text = MyData.dtmainMuqavileRekvizit.Rows[0]["Vöen qeydiyyat ünvanı"].ToString(); }
                            catch { /*txtunvan.Text = ""; */}
                            try { txtSurvesiqe1.Text = MyData.dtmainMuqavileRekvizit.Rows[0]["Vöen"].ToString(); }
                            catch { /*txtunvan.Text = ""; */}
                            try { txtsexsiyyet1.Text = MyData.dtmainMuqavileRekvizit.Rows[0]["Direktor"].ToString(); }
                            catch { /*txtunvan.Text = ""; */}
                        }
                        else { txtunvan1.Text = ""; tel1.Text = ""; txtSurvesiqe1.Text = ""; txtsexsiyyet1.Text = ""; }
                    }
                }
                catch { }
            }
        }

        private void txtsurucu_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    txtsurucu.Text = txtsurucu.Text.Substring(0, 1).ToUpper(MyChange.DilDeyisme) + txtsurucu.Text.Substring(1, txtsurucu.Text.Length - 1).ToLower(MyChange.DilDeyisme);
                }
                catch { }

                try
                {
                    MyData.selectCommand("baza.accdb", "Select * from etibarnamesurucu where a1 Like '%" + txtsurucu.Text + "%'");
                    MyData.dtmainSurucu = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmainSurucu);

                    if (MyData.dtmainSurucu.Rows.Count > 0 & txtsurucu.Text != "")
                    {
                        for (int i = 0; i < MyData.dtmainSurucu.Rows.Count; i++)
                        {
                            try { txtsurucu.Text = MyData.dtmainSurucu.Rows[i]["a1"].ToString(); }
                            catch { /*t7.Text = "";*/ }
                            try { tel2.Text = MyData.dtmainSurucu.Rows[i]["a8"].ToString(); }
                            catch { /*tel1.Text = "";*/ }
                            try { txtunvan2.Text = MyData.dtmainSurucu.Rows[i]["a4"].ToString(); }
                            catch { /*txtunvan.Text = ""; */}
                            try { txtSurvesiqe2.Text = MyData.dtmainSurucu.Rows[i]["a3"].ToString(); }
                            catch { /*txtunvan.Text = ""; */}
                            try { txtsexsiyyet2.Text = MyData.dtmainSurucu.Rows[i]["a2"].ToString(); }
                            catch { /*txtunvan.Text = ""; */}
                        }
                    }
                    else { txtunvan2.Text = ""; tel2.Text = ""; txtSurvesiqe2.Text = ""; txtsexsiyyet2.Text = ""; }

                }
                catch { }
            }
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            try { txtsurucu.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Column13"].Value.ToString(); }
            catch { }

            ///////////////////Surucunun melumatlarinin axtarisi
             try
             {
                 MyData.selectCommand("baza.accdb", "Select * from etibarnamesurucu where a1 Like '%" + txtsurucu.Text + "%'");
                 MyData.dtmainSurucu = new DataTable();
                 MyData.oledbadapter1.Fill(MyData.dtmainSurucu);

                 if (MyData.dtmainSurucu.Rows.Count > 0 & txtsurucu.Text != "")
                 {
                     for (int i = 0; i < MyData.dtmainSurucu.Rows.Count; i++)
                     {
                         try { txtsurucu.Text = MyData.dtmainSurucu.Rows[i]["a1"].ToString(); }
                         catch { /*t7.Text = "";*/ }
                         try { tel2.Text = MyData.dtmainSurucu.Rows[i]["a8"].ToString(); }
                         catch { /*tel1.Text = "";*/ }
                         try { txtunvan2.Text = MyData.dtmainSurucu.Rows[i]["a4"].ToString(); }
                         catch { /*txtunvan.Text = ""; */}
                         try { txtSurvesiqe2.Text = MyData.dtmainSurucu.Rows[i]["a3"].ToString(); }
                         catch { /*txtunvan.Text = ""; */}
                         try { txtsexsiyyet2.Text = MyData.dtmainSurucu.Rows[i]["a2"].ToString(); }
                         catch { /*txtunvan.Text = ""; */}
                     }
                 }
                 else { txtunvan2.Text = ""; tel2.Text = ""; txtSurvesiqe2.Text = ""; txtsexsiyyet2.Text = ""; }

             }
             catch { }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Metn = "";

            Metn = "Salam," + Environment.NewLine + Environment.NewLine + comboBox1.Text + " - yə məxsus dövlət nömrə nişanı " + t2.Text + " olan " + t3.Text + " markalı nəqliyyat vasitəsi " + t5.Text + " - cu il tarixdə saat " + dateTimePicker1.Text + " - də " + t1.Text + " üzrə inzibati xəta törətmişdir. Qeyd olunan nəqliyyat vasitəsi üzrə məlumatları sizə təqdim edirik." + Environment.NewLine + Environment.NewLine;
            if (checkBox1.Checked == true) Metn += "Lizinq alan:" + Environment.NewLine + txtlizinqalan.Text + Environment.NewLine + lbSur.Text + ": " + txtSurvesiqe1.Text + Environment.NewLine + lbSexsiyyet.Text + ": " + txtsexsiyyet1.Text + Environment.NewLine + "Ünvan: " + txtunvan1.Text + Environment.NewLine + "Tel: " + tel1.Text + Environment.NewLine + Environment.NewLine;
            if (checkBox2.Checked == true) Metn += "Sürücü:" + Environment.NewLine + txtsurucu.Text + Environment.NewLine + "Sürücüluk vəsiqəsi: " + txtSurvesiqe2.Text + Environment.NewLine + "Ş/V: " + txtsexsiyyet2.Text + Environment.NewLine + "Ünvan: " + txtunvan2.Text + Environment.NewLine + "Tel: " + tel2.Text + Environment.NewLine + Environment.NewLine;

            try { File.Copy("New Emphty.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\BNA Sorğu - " + txtSorgu.Text + " - " + t2.Text + ".doc", true); }
            catch { MessageBox.Show("'\\192.168.10.5\\Common\\AGLizinq\\New Emphty.doc' tapılmadı."); }

            object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\BNA Sorğu - " + txtSorgu.Text + " - " + t2.Text + ".doc";

            Microsoft.Office.Interop.Word._Application oWord = new Microsoft.Office.Interop.Word.Application();
            object oMissing = Type.Missing;
            oWord.Visible = true;
            oWord.Documents.Open(FileName);
            oWord.Selection.TypeText(Metn);
            oWord.ActiveDocument.Save();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            SendMail();
        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            try { txtsurucu.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["a1"].Value.ToString(); }
            catch { /*t7.Text = "";*/ }
            try { tel2.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["a8"].Value.ToString(); }
            catch { /*tel1.Text = "";*/ }
            try { txtunvan2.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["a4"].Value.ToString(); }
            catch { /*txtunvan.Text = ""; */}
            try { txtSurvesiqe2.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["aa3"].Value.ToString(); }
            catch { /*txtunvan.Text = ""; */}
            try { txtsexsiyyet2.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["a2"].Value.ToString(); }
            catch { /*txtunvan.Text = ""; */}
        }

        private void BNAadd_Load(object sender, EventArgs e)
        {

        }
    }
}
