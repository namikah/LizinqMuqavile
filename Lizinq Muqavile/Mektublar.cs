using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Nsoft;

namespace Lizinq_Muqavile
{
    public partial class Mektublar : Form
    {

        Excel.Application oXL;
        Excel._Workbook oWB;
        Excel._Worksheet oSheet;

        public Mektublar()
        {
            InitializeComponent();
        }
        public int emeliyyatUcun = 0;

        public void Mektub_itirilme()
        {
            try { File.Copy("Mektublar\\Mektub-itirilme.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Mektub-itirilme (" + txtnomre.Text + ").doc", true); }
            catch { MessageBox.Show("'\\192.168.10.5\\Common\\AGLizinq\\Mektub-itirilme.doc' tapılmadı."); }

                object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\Mektub-itirilme (" + txtnomre.Text + ").doc";


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

            if (comboBox1.Text == "«AGLizinq» QSC")
            {
                MyChange.FindAndReplace(word, "000", "№ 000/18");
                MyChange.FindAndReplace(word, "111", dttarix.Text.Substring(0, 2) + " " + b + " " + dttarix.Text.Substring(dttarix.Text.Length - 4, 4) + " - ci il");
            }
            else
            {
                MyChange.FindAndReplace(word, "000", "");
                MyChange.FindAndReplace(word, "111", "");
            }

            MyChange.FindAndReplace(word, "000000", comboBox2.Text);
            MyChange.FindAndReplace(word, "222", comboBox1.Text);
            MyChange.FindAndReplace(word, "333", txtmarka.Text);
            MyChange.FindAndReplace(word, "333", txtmarka.Text);
            MyChange.FindAndReplace(word, "444", txtnomre.Text);
            MyChange.FindAndReplace(word, "555", txtburaxilis.Text);
            MyChange.FindAndReplace(word, "666", txtsehadetname.Text);
            MyChange.FindAndReplace(word, "777", txtmuherrik.Text);
            MyChange.FindAndReplace(word, "888", txtban.Text);
            MyChange.FindAndReplace(word, "999", txtzavodnomresi.Text);
            MyChange.FindAndReplace(word, "10000", txtrengi.Text);
            MyChange.FindAndReplace(word, "20000", comboBox4.Text);
            MyChange.FindAndReplace(word, "200002", comboBox3.Text);
            
            doc.Save();

        }

        public void Erize_DYP()
        {
            try
            {
                File.Copy("Mulkiyyete verme\\Erize.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Erize - " + txtnomre.Text + ".xlsx", true);
            }
            catch { MessageBox.Show("Erize.xlsx tapılmadı."); }

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Erize - " + txtnomre.Text + ".xlsx"));
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

        private void button2_Click(object sender, EventArgs e)
        {
            if (comboBox5.Text == "MƏKTUB")
            {
                Mektub_itirilme();
            }

            if (comboBox5.Text == "ƏRİZƏ") 
            { 
                Erize_DYP(); 
            }
        }

        private void txtnomre_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    MyData.selectCommand("baza.accdb","Select * from etibarnameneqliyyat where c1 Like '%" + txtnomre.Text + "%'");
                    MyData.dtmain = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmain);

                    txtmarka.Text = MyData.dtmain.Rows[0]["c2"].ToString(); txtmarka.BackColor = Color.LightGreen;
                    txtnomre.Text = MyData.dtmain.Rows[0]["c1"].ToString(); txtnomre.BackColor = Color.LightGreen;
                    txtsehadetname.Text = MyData.dtmain.Rows[0]["c12"].ToString(); txtsehadetname.BackColor = Color.LightGreen;
                    txtsassi.Text = MyData.dtmain.Rows[0]["c10"].ToString(); txtsassi.BackColor = Color.LightGreen;
                    txtban.Text = MyData.dtmain.Rows[0]["c8"].ToString(); txtban.BackColor = Color.LightGreen;
                    txtmuherrik.Text = MyData.dtmain.Rows[0]["c9"].ToString(); txtmuherrik.BackColor = Color.LightGreen;
                    txtrengi.Text = MyData.dtmain.Rows[0]["c5"].ToString(); txtrengi.BackColor = Color.LightGreen;
                    txtburaxilis.Text = MyData.dtmain.Rows[0]["c6"].ToString(); txtburaxilis.BackColor = Color.LightGreen;
                    txtzavodnomresi.Text = MyData.dtmain.Rows[0]["c11"].ToString(); txtzavodnomresi.BackColor = Color.LightGreen;
                    txtlizinqalan.Text = MyData.dtmain.Rows[0]["c3"].ToString(); txtlizinqalan.BackColor = Color.LightGreen;
                }
                catch { }

                MyData.selectCommand("baza.accdb", "Select * from Telefon where c1 Like " + "'%" + txtlizinqalan.Text + "%'");
                MyData.dtmain = new DataTable();
               MyData.oledbadapter1.Fill(MyData.dtmain);

                tel1.Text = ""; tel2.Text = "";
                try { tel1.Text = MyData.dtmain.Rows[0]["c2"].ToString(); tel1.BackColor = Color.LightGreen; }
                catch { };
                try { tel2.Text = MyData.dtmain.Rows[0]["c3"].ToString(); tel1.BackColor = Color.LightGreen; }
                catch { };

            }
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox3.Items.Clear(); comboBox3.Text = ""; comboBox3.Enabled = false;

            if (comboBox4.Text == "dövlət qeydiyyat nişanları naməlum şəraitdə itirilmişdir.")
            {
                //comboBox3.Text = "qeydiyyat nişanının tapılmasında Bizə köməklik göstərəsiniz.";
                comboBox3.Items.Add("yeni-təkrar qeydiyyat nişanı verəsiniz.");
                comboBox3.Items.Add("qeydiyyat nişanının tapılmasında Bizə köməklik göstərəsiniz."); 
                comboBox3.Enabled = true;
            }

            if (comboBox4.Text == "qeydiyyat şəhadətnaməsi naməlum şəraitdə itirilmişdir.")
            {
                //comboBox3.Text = "qeydiyyat şəhadətnaməsinin tapılmasında Bizə köməklik göstərəsiniz.";
                comboBox3.Items.Add("yeni-təkrar qeydiyyat şəhadətnaməsi verəsiniz.");
                comboBox3.Items.Add("qeydiyyat şəhadətnaməsinin tapılmasında Bizə köməklik göstərəsiniz.");
                comboBox3.Enabled = true;
            }

            if (comboBox4.Text == "dövlət qeydiyyat nişanları və qeydiyyat şəhadətnaməsi naməlum şəraitdə itirilmişdir.")
            {
                //comboBox3.Text = "qeydiyyat şəhadətnaməsinin və qeydiyyat nişanlarının tapılmasında Bizə köməklik göstərəsiniz.";
                comboBox3.Items.Add("yeni-təkrar qeydiyyat şəhadətnaməsi və qeydiyyat nişanları verəsiniz.");
                comboBox3.Items.Add("qeydiyyat şəhadətnaməsinin və qeydiyyat nişanlarının tapılmasında Bizə köməklik göstərəsiniz.");
                comboBox3.Enabled = true;
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox3.Items.Clear(); comboBox3.Text = ""; comboBox3.Enabled = false;
            comboBox4.Items.Clear(); comboBox4.Text = ""; comboBox4.Enabled = false;

            if (comboBox2.Text == "AR DİN BAKI ŞƏHƏR BAŞ POLİS İDARƏSİNİN TAPINTILAR BÜROSUNA")
            {
                //comboBox4.Text = "dövlət qeydiyyat nişanları naməlum şəraitdə itirilmişdir.";
                comboBox4.Items.Add("dövlət qeydiyyat nişanları naməlum şəraitdə itirilmişdir.");
                comboBox4.Items.Add("qeydiyyat şəhadətnaməsi naməlum şəraitdə itirilmişdir.");
                comboBox4.Items.Add("dövlət qeydiyyat nişanları və qeydiyyat şəhadətnaməsi naməlum şəraitdə itirilmişdir.");
                comboBox4.Enabled = true;
            }

            else if (comboBox2.Text == "AZƏRBAYCAN RESPUBLİKASI DAXİLİ İŞLƏR NAZİRLİYİ BAŞ DÖVLƏT YOL POLİSİ  İDARƏSİNİN QEYDİYYAT-İMTAHAN ŞÖBƏSİNƏ")
            {
                //comboBox4.Text = "dövlət qeydiyyat nişanları naməlum şəraitdə itirilmişdir.";
                comboBox4.Items.Add("dövlət qeydiyyat nişanları naməlum şəraitdə itirilmişdir.");
                comboBox4.Items.Add("qeydiyyat şəhadətnaməsi naməlum şəraitdə itirilmişdir.");
                comboBox4.Items.Add("dövlət qeydiyyat nişanları və qeydiyyat şəhadətnaməsi naməlum şəraitdə itirilmişdir.");
                 comboBox4.Enabled = true;
            }

            else if (comboBox2.Text == "DİN BAŞ DÖVLƏT YOL POLİS İDARƏSİNİN RƏİS MÜAVİNİ Qİİ-NİN RƏİSİ POLİS POLKOVNİKİ CƏNAB M.ŞAHBAZOVA")
            {
                //comboBox4.Text = "dövlət qeydiyyat nişanları naməlum şəraitdə itirilmişdir.";
                comboBox4.Items.Add("dövlət qeydiyyat nişanları naməlum şəraitdə itirilmişdir.");
                comboBox4.Items.Add("qeydiyyat şəhadətnaməsi naməlum şəraitdə itirilmişdir.");
                comboBox4.Items.Add("dövlət qeydiyyat nişanları və qeydiyyat şəhadətnaməsi naməlum şəraitdə itirilmişdir.");
                 comboBox4.Enabled = true;
            }

            else if (comboBox2.Text == "YASAMAL RAYONU 26 SAYLI POLİS BÖLMƏSİNƏ")
            {
                //comboBox4.Text = "dövlət qeydiyyat nişanları naməlum şəraitdə itirilmişdir.";
                comboBox4.Items.Add("dövlət qeydiyyat nişanları naməlum şəraitdə itirilmişdir.");
                comboBox4.Items.Add("qeydiyyat şəhadətnaməsi naməlum şəraitdə itirilmişdir.");
                comboBox4.Items.Add("dövlət qeydiyyat nişanları və qeydiyyat şəhadətnaməsi naməlum şəraitdə itirilmişdir.");
                comboBox4.Enabled = true;
            }
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox3.Items.Clear(); comboBox3.Text = ""; comboBox3.Enabled = false;
            comboBox4.Items.Clear(); comboBox4.Text = ""; comboBox4.Enabled = false;
            comboBox2.Items.Clear(); comboBox2.Text = ""; comboBox2.Enabled = false;
           

            if (comboBox5.Text == "MƏKTUB")
            {
                comboBox2.Items.Add("AR DİN BAKI ŞƏHƏR BAŞ POLİS İDARƏSİNİN TAPINTILAR BÜROSUNA");
                comboBox2.Items.Add("AZƏRBAYCAN RESPUBLİKASI DAXİLİ İŞLƏR NAZİRLİYİ BAŞ DÖVLƏT YOL POLİSİ  İDARƏSİNİN QEYDİYYAT-İMTAHAN ŞÖBƏSİNƏ");
                comboBox2.Items.Add("DİN BAŞ DÖVLƏT YOL POLİS İDARƏSİNİN RƏİS MÜAVİNİ Qİİ-NİN RƏİSİ POLİS POLKOVNİKİ CƏNAB M.ŞAHBAZOVA");
                comboBox2.Items.Add("YASAMAL RAYONU 26 SAYLI POLİS BÖLMƏSİNƏ");
                comboBox2.Enabled = true;
                button2.BackgroundImage = Lizinq_Muqavile.Properties.Resources.word;

                label20.Text = "vacib deyil";
                label21.Text = "vacib deyil";
                label22.Text = "vacib deyil";
                label23.Text = "vacib deyil";
                label24.Text = "vacib deyil";
                label25.Text = "vacib deyil";
                label26.Text = "vacib deyil";
            }

            else if (comboBox5.Text == "ƏRİZƏ")
            {
                comboBox2.Text = "AZƏRBAYCAN RESPUBLİKASI DAXİLİ İŞLƏR NAZİRLİYİ BAŞ DÖVLƏT YOL POLİSİ İDARƏSİNƏ";
                comboBox2.Items.Add("AZƏRBAYCAN RESPUBLİKASI DAXİLİ İŞLƏR NAZİRLİYİ BAŞ DÖVLƏT YOL POLİSİ İDARƏSİNƏ");
                comboBox2.Enabled = true;
                button2.BackgroundImage = Lizinq_Muqavile.Properties.Resources.excel;

                label20.Text = "*";
                label21.Text = "*";
                label22.Text = "*";
                label23.Text = "*";
                label24.Text = "*";
                label25.Text = "*";
                label26.Text = "*"; 
                label28.Text = "*"; 
                label29.Text = "*";
            }
        }

        private void Mektublar_Load(object sender, EventArgs e)
        {

        }
    }
}
