using Nsoft;
using System;
using System.Data;
using System.IO;
using System.Net.Mail;
using System.Windows.Forms;

namespace Lizinq_Muqavile
{
    public partial class MMXmelumat : Form
    {
        Excel.Application oXL;
        Excel._Workbook oWB;
        Excel._Worksheet oSheet;

        public MMXmelumat()
        {
            InitializeComponent();
        }

        private void MMXMelumatNomre()
        {
            MyData.selectCommand("baza.accdb", "Select * from MMXMelumatNomre");
            MyData.dtmainNomre= new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainNomre);

            txtNomre.Text = (Convert.ToInt32(MyData.dtmainNomre.Rows[0][0]) + 1).ToString();

            try
            {
                if (Convert.ToInt32(MyData.dtmainNomre.Rows[0][0]) > 0 && Convert.ToInt32(MyData.dtmainNomre.Rows[0][0]) < 10) txtNomre.Text = $"00000{(Convert.ToInt32(MyData.dtmainNomre.Rows[0][0]) + 1).ToString()}";


                if (Convert.ToInt32(MyData.dtmainNomre.Rows[0][0]) > 9 && Convert.ToInt32(MyData.dtmainNomre.Rows[0][0]) < 100) txtNomre.Text = $"0000{(Convert.ToInt32(MyData.dtmainNomre.Rows[0][0]) + 1).ToString()}";


                if (Convert.ToInt32(MyData.dtmainNomre.Rows[0][0]) > 99 && Convert.ToInt32(MyData.dtmainNomre.Rows[0][0]) < 1000) txtNomre.Text = $"000{(Convert.ToInt32(MyData.dtmainNomre.Rows[0][0]) + 1).ToString()}";


                if (Convert.ToInt32(MyData.dtmainNomre.Rows[0][0]) > 999 && Convert.ToInt32(MyData.dtmainNomre.Rows[0][0]) < 10000) txtNomre.Text = $"00{(Convert.ToInt32(MyData.dtmainNomre.Rows[0][0]) + 1).ToString()}";


                if (Convert.ToInt32(MyData.dtmainNomre.Rows[0][0]) > 9999 && Convert.ToInt32(MyData.dtmainNomre.Rows[0][0]) < 100000) txtNomre.Text = $"0{(Convert.ToInt32(MyData.dtmainNomre.Rows[0][0]) + 1).ToString()}";


                if (Convert.ToInt32(MyData.dtmainNomre.Rows[0][0]) > 99999 && Convert.ToInt32(MyData.dtmainNomre.Rows[0][0]) < 1000000) txtNomre.Text = (Convert.ToInt32(MyData.dtmainNomre.Rows[0][0]) + 1).ToString();

            }
            catch { };

        }

        private void EtibarnameCap()
        {
            try
            {
                Directory.CreateDirectory("X:\\Umumi Senedler\\MMX\\" + t2.Text + "\\"); File.Copy("Etibarname.xlsx", "X:\\Umumi Senedler\\MMX\\" + t2.Text + "\\" + t2.Text + " (" + t5.Text + ") Etibarname.xlsx", true);
            }
            catch { }

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("X:\\Umumi Senedler\\MMX\\" + t2.Text + "\\" + t2.Text + " (" + t5.Text + ") Etibarname.xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];

            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            try
            {
                oSheet.Cells[1, 4] = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString();
                oSheet.Cells[3, 3] = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[2].Value.ToString();
                oSheet.Cells[5, 2] = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[3].Value.ToString();
                oSheet.Cells[3, 7] = "Şəhadətnaməsi əsasında " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[3].Value.ToString() + "-yə məxsus";
            }
            catch { MessageBox.Show("Arxivdə yazılmış etibarnamə yoxdur."); return; }

            if (comboBox1.Text.Substring(1, 3) == "AGB") { oSheet.Cells[7, 2] = "AZ1022 Bakı şəh, Nəsimi r-nu,"; oSheet.Cells[8, 2] = "Cəlil Məmmədquluzadə, ev 102A"; }
            else { oSheet.Cells[7, 2] = "AZ1073 Bakı şəh, Yasamal r-nu,"; oSheet.Cells[8, 2] = "Landau küç 16"; }

            oSheet.Cells[9, 2] = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[4].Value.ToString();
            if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Value.ToString() != "")
            {
                oSheet.Cells[12, 1] = "Şəxsiyyət vəsiqəsi:";
                oSheet.Cells[12, 4] = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Value.ToString();
            }
            if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[6].Value.ToString() != "")
            {
                oSheet.Cells[12, 1] = "Sürücülük vəsiqəsi:";
                oSheet.Cells[12, 4] = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[6].Value.ToString();
            }
            oSheet.Cells[14, 1] = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[7].Value.ToString();
            oSheet.Cells[18, 1] = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[8].Value.ToString();
            oSheet.Cells[23, 1] = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[9].Value.ToString();


            oSheet.Cells[1, 7] = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[10].Value.ToString();
            oSheet.Cells[5, 7] = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[11].Value.ToString();
            oSheet.Cells[7, 10] = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[12].Value.ToString();

            try { if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[13].Value.ToString() != "") { oSheet.Cells[8, 8] = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[13].Value.ToString(); } }
            catch { }
            try { if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[14].Value.ToString() != "") { oSheet.Cells[8, 8] = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[13].Value.ToString(); } }
            catch { }

            oSheet.Cells[9, 8] = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[15].Value.ToString();
            oSheet.Cells[11, 8] = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[16].Value.ToString();
            oSheet.Cells[20, 11] = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[17].Value.ToString();
            oSheet.Cells[19, 9] = "";

            oSheet.Unprotect();
            try { oSheet.Cells[26, 1] = txtsurucu.Text; }
            catch { }
            try { oSheet.Cells[27, 1] = "Mob: " + tel1.Text + " / " + tel2.Text; }
            catch { }

            //  oSheet.PrintOut();
            //   oWB.Close(SaveChanges: false);
            // oXL.Application.Quit();

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
            //oSheet.PrintOut();
            //oWB.Close(SaveChanges: false);
            //oXL.Application.Quit();
        }

        private void MelumatSurucu()
        {
            try { Directory.CreateDirectory("X:\\Umumi Senedler\\MMX\\" + t2.Text + "\\"); File.Copy("MMX\\MMX Melumat.xlsx", "X:\\Umumi Senedler\\MMX\\" + t2.Text + "\\" + t2.Text + " (" + t5.Text + ") Surucu Melumat.xlsx", true); }
            catch { MessageBox.Show("'\\192.168.10.5\\Common\\AGLizinq\\Hesabat Lizinq.xlsx' tapılmadı."); }

            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("X:\\Umumi Senedler\\MMX\\" + t2.Text + "\\" + t2.Text + " (" + t5.Text + ") Surucu Melumat.xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];

            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            
            
           
            MyData.selectCommand("baza.accdb", "Select * from etibarnamesurucu where a1 Like '%" + txtsurucu.Text + "%'");
            MyData.dtmainSurucuMelumat= new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainSurucuMelumat);

            try { oSheet.Cells[4, 1] = "BAKI ŞƏHƏRİ " + comboBox2.Text + " RAYON POLİS İDARƏSİNƏ"; }
            catch { }

            string l = "", s = "";

            if (checkBox1.Checked == true)
            {
                l = txtlizinqalan.Text + Environment.NewLine + lbSur.Text + ": " + txtSurvesiqe1.Text + Environment.NewLine + lbSexsiyyet.Text + ": " + txtsexsiyyet1.Text + Environment.NewLine + "Ünvan: " + txtunvan1.Text + Environment.NewLine + "Tel: " + tel1.Text;
            }

            if (checkBox2.Checked == true)
            {
                s = txtsurucu.Text + Environment.NewLine + lbSur2.Text + ": " + txtSurvesiqe2.Text + Environment.NewLine + lbSexsiyyet2.Text + ": " + txtsexsiyyet2.Text + Environment.NewLine + "Ünvan: " + txtunvan2.Text + Environment.NewLine + "Tel: " + tel2.Text;
            }
            else
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    s+= dataGridView1.Rows[i].Cells["Column11"].Value.ToString() + " - " + dataGridView1.Rows[i].Cells["Column13"].Value.ToString() + " - " +dataGridView1.Rows[i].Cells["Column7"].Value.ToString() + " - " +dataGridView1.Rows[i].Cells["Column22"].Value.ToString() + Environment.NewLine;
                }
            }

            oSheet.Cells[9, 1] = Convert.ToInt32(1).ToString();
            oSheet.Cells[9, 2] = comboBox1.Text;
            oSheet.Cells[9, 3] = t1.Text;
            oSheet.Cells[9, 4] = t3.Text;
            oSheet.Cells[9, 5] = t2.Text;
            oSheet.Cells[9, 6] = l.ToString();
            oSheet.Cells[9, 7] = s.ToString();

            oSheet.Range["A" + 9].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["B" + 9].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["C" + 9].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["D" + 9].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["E" + 9].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["F" + 9].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["G" + 9].Borders.LineStyle = Excel.Constants.xlSolid;

            //  oSheet.PrintOut();
            //   oWB.Close(SaveChanges: false);
            // oXL.Application.Quit();

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
            oSheet.Columns.AutoFit();
            oSheet.Rows.AutoFit();
            //oSheet.PrintOut();
            //oWB.Close(SaveChanges: false);
            //oXL.Application.Quit();
        }

        public void MektubAGBank()
        {
            try { Directory.CreateDirectory("X:\\Umumi Senedler\\MMX\\" + t2.Text + "\\"); File.Copy("MMX\\MMX Mektub AGBank.doc", "X:\\Umumi Senedler\\MMX\\" + t2.Text + "\\" + t2.Text + " (" + t5.Text + ") Mektub.doc", true); }
            catch { MessageBox.Show("'\\192.168.10.5\\Common\\AGLizinq\\MMX Mektub.doc' tapılmadı."); }

            object FileName = "X:\\Umumi Senedler\\MMX\\" + t2.Text + "\\" + t2.Text + " (" + t5.Text + ") Mektub.doc";

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

            string a = MyChange.TarixSozle(Convert.ToDateTime(t4.Text));
            string b = MyChange.TarixSozle(Convert.ToDateTime(t5.Text));
            string c = MyChange.TarixSozle(DateTime.Today);

            MyChange.FindAndReplace(word, "000", txtNomre.Text);
            MyChange.FindAndReplace(word, "1111", "AZƏRBAYCAN RESPUBLİKASI DAXİLİ İŞLƏR NAZİRLİYİ" + Environment.NewLine + "BAKI ŞƏHƏRİ " + comboBox2.Text + " RAYON POLİS İDARƏSİNİN" + Environment.NewLine + "DÖVLƏT YOL POLİSİ ŞÖBƏSİNƏ");
            MyChange.FindAndReplace(word, "111", DateTime.Today.ToShortDateString().Substring(0, 2) + " " + b + " " + DateTime.Today.ToShortDateString().Substring(6, 4) + " - ci il");
            MyChange.FindAndReplace(word, "222", t4.Text.Substring(0, 2) + " " + a + " " + t4.Text.Substring(6, 4) + " - ci il");
            MyChange.FindAndReplace(word, "333", t2.Text);
            MyChange.FindAndReplace(word, "444", t3.Text);
            MyChange.FindAndReplace(word, "555", t5.Text.Substring(0, 2) + " " + c + " " + t5.Text.Substring(6, 4) + " - ci il");
            MyChange.FindAndReplace(word, "666", t1.Text);

            if (checkBox5.Checked == true)
            {
                MyChange.FindAndReplace(word, "777", "elektron protokol tərtib edildiyi dövr üçün sürücülük hüququ (etibarnamə) verilən şəxslərin etibarnamələrinin elektron çıxarışını Sizə təqdim edirik. ");
            }
            else { MyChange.FindAndReplace(word, "777", " məlumatları Sizə təqdim edirik."); }

            doc.Save();
        }

        public void MektubAGLizinq()
        {
            try { Directory.CreateDirectory("X:\\Umumi Senedler\\MMX\\" + t2.Text + "\\"); File.Copy("MMX\\MMX Mektub AGLizinq.doc", "X:\\Umumi Senedler\\MMX\\" + t2.Text + "\\" + t2.Text + " (" + t5.Text + ") Mektub.doc", true); }
            catch { MessageBox.Show("'\\192.168.10.5\\Common\\AGLizinq\\MMX Melumat.doc' tapılmadı."); }

            object FileName = "X:\\Umumi Senedler\\MMX\\" + t2.Text + "\\" + t2.Text + " (" + t5.Text + ") Mektub.doc";

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

            string a = MyChange.TarixSozle(Convert.ToDateTime(t4.Text));
            string b = MyChange.TarixSozle(Convert.ToDateTime(t5.Text));
            string c = MyChange.TarixSozle(DateTime.Today);

            MyChange.FindAndReplace(word, "000", txtNomre.Text);
            MyChange.FindAndReplace(word, "1111", "AZƏRBAYCAN RESPUBLİKASI DAXİLİ İŞLƏR NAZİRLİYİ" + Environment.NewLine + "BAKI ŞƏHƏRİ " + comboBox2.Text + " RAYON POLİS İDARƏSİNİN" + Environment.NewLine + "DÖVLƏT YOL POLİSİ ŞÖBƏSİNƏ");
            MyChange.FindAndReplace(word, "111", DateTime.Today.ToShortDateString().Substring(0, 2) + " " + b + " " + DateTime.Today.ToShortDateString().Substring(6, 4) + " - ci il");
            MyChange.FindAndReplace(word, "222", t4.Text.Substring(0, 2) + " " + a + " " + t4.Text.Substring(6, 4) + " - ci il");
            MyChange.FindAndReplace(word, "333", t2.Text);
            MyChange.FindAndReplace(word, "444", t3.Text);
            MyChange.FindAndReplace(word, "555", t5.Text.Substring(0, 2) + " " + c + " " + t5.Text.Substring(6, 4) + " - ci il");
            MyChange.FindAndReplace(word, "666", t1.Text);

            if (checkBox5.Checked == true)
            {
                MyChange.FindAndReplace(word, "777", "elektron protokol tərtib edildiyi dövr üçün sürücülük hüququ (etibarnamə) verilən şəxslərin etibarnamələrinin elektron çıxarışını Sizə təqdim edirik. ");
            }
            else { MyChange.FindAndReplace(word, "777", " məlumatları Sizə təqdim edirik."); }

            doc.Save();
        }

        public void BildirisAGBank()
        {
            try { Directory.CreateDirectory("X:\\Umumi Senedler\\MMX\\" + t2.Text + "\\"); File.Copy("MMX\\Bildiris AGBank.doc", "X:\\Umumi Senedler\\MMX\\" + t2.Text + "\\" + txtNomre.Text + " protokol Bildiris " + t2.Text + " (" + t5.Text + ").doc", true); }
            catch { MessageBox.Show("'\\192.168.10.5\\Common\\AGLizinq\\MMX Melumat.doc' tapılmadı."); }

            object FileName = "X:\\Umumi Senedler\\MMX\\" + t2.Text + "\\" + txtNomre.Text + " protokol Bildiris " + t2.Text + " (" + t5.Text + ").doc";

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

            string a = MyChange.TarixSozle(Convert.ToDateTime(t4.Text));

            MyChange.FindAndReplace(word, "000", txtNomre.Text);
            MyChange.FindAndReplace(word, "111", DateTime.Today.ToShortDateString().Substring(0, 2) + " " + a + " " + DateTime.Today.ToShortDateString().Substring(6, 4) + " - ci il");
            MyChange.FindAndReplace(word, "222", txtsurucu.Text);
            MyChange.FindAndReplace(word, "333", txtunvan2.Text);
            MyChange.FindAndReplace(word, "444", t2.Text);
            MyChange.FindAndReplace(word, "555", t1.Text);

            doc.Save();
        }

        public void BildirisAGLizinq()
        {
            try { Directory.CreateDirectory("X:\\Umumi Senedler\\MMX\\" + t2.Text + "\\"); File.Copy("MMX\\Bildiris AGLizinq.doc", "X:\\Umumi Senedler\\MMX\\" + t2.Text + "\\" + txtNomre.Text + " protokol Bildiris " + t2.Text + " (" + t5.Text + ").doc", true); }
            catch { MessageBox.Show("'\\192.168.10.5\\Common\\AGLizinq\\MMX Melumat.doc' tapılmadı."); }

            object FileName = "X:\\Umumi Senedler\\MMX\\" + t2.Text + "\\" + txtNomre.Text + " protokol Bildiris " + t2.Text + " (" + t5.Text + ").doc";

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

            string a = MyChange.TarixSozle(Convert.ToDateTime(t4.Text));

            MyChange.FindAndReplace(word, "000", txtNomre.Text);
            MyChange.FindAndReplace(word, "111", DateTime.Today.ToShortDateString().Substring(0, 2) + " " + a + " " + DateTime.Today.ToShortDateString().Substring(6, 4) + " - ci il");
            MyChange.FindAndReplace(word, "222", txtsurucu.Text);
            MyChange.FindAndReplace(word, "333", txtunvan2.Text);
            MyChange.FindAndReplace(word, "444", t2.Text);
            MyChange.FindAndReplace(word, "555", t1.Text);

            doc.Save();
        }

        private void SendMail()
        {
            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient("data4.web.az");
                //SmtpClient SmtpServer = new SmtpClient("Smtp.mail.ru");
                //mail.From = new MailAddress("aglizinq@mail.ru");

                mail.From = new MailAddress("info@agleasing.az");
                if (txtMail.Text != "namikah@agleasing.az" && checkBox3.Checked == true) mail.To.Add("namikah@agleasing.az");
                if (txtMail.Text != "rashadim@agleasing.az" && checkBox3.Checked == true) mail.To.Add("rashadim@agleasing.az");
                try { if (checkBox4.Checked == true) mail.To.Add(txtMail.Text); }
                catch { }
                //mail.To.Add("heydarovnamik@mail.ru");
                mail.Subject = "MMX (" + t2.Text + ")";
                mail.Body = label1.Text + " - " + t1.Text + Environment.NewLine + Environment.NewLine + label22.Text + " - " + t2.Text + Environment.NewLine + Environment.NewLine + label21.Text + " - " + t3.Text + Environment.NewLine + Environment.NewLine + label20.Text + " - " + t4.Text + Environment.NewLine + Environment.NewLine + label6.Text + " - " + t5.Text + Environment.NewLine + Environment.NewLine + label18.Text + " - " + t6.Text + Environment.NewLine + Environment.NewLine + label17.Text + " - " + txtsurucu.Text + Environment.NewLine + Environment.NewLine + "Cərimə məbləği - " + t10.Text + " AZN";

                SmtpServer.Port = 587;
                SmtpServer.Credentials = new System.Net.NetworkCredential("info@agleasing.az", "Lhx-9.9-lhx");
                //SmtpServer.Credentials = new System.Net.NetworkCredential("aglizinq@mail.ru", "Lhx99lhx");
                SmtpServer.EnableSsl = true;
                SmtpServer.Send(mail);
            }
            catch { }
        }

        private void t2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {

                try
                {
                    t2.Text = t2.Text.Substring(0, 1).ToUpper(MyChange.DilDeyisme) + t2.Text.Substring(1, t2.Text.Length - 1).ToLower(MyChange.DilDeyisme);
                }
                catch { }

                try
                {
                    
                    
                   
                    MyData.selectCommand("baza.accdb", "Select * from etibarnameneqliyyat where c1 Like " + "'%" + t2.Text + "%'");
                    MyData.dtmainEtibarbame= new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmainEtibarbame);

                    if (MyData.dtmainEtibarbame.Rows[0]["c13"].ToString() == "AGB") comboBox1.Text = "“AGBank” ASC";
                    if (MyData.dtmainEtibarbame.Rows[0]["c13"].ToString() == "AGL") comboBox1.Text = "“AGLizinq” QSC";

                    t2.Text = MyData.dtmainEtibarbame.Rows[0]["c1"].ToString();
                    t3.Text = MyData.dtmainEtibarbame.Rows[0]["c2"].ToString();
                    txtlizinqalan.Text = MyData.dtmainEtibarbame.Rows[0]["c3"].ToString();
                    txtLayihe.Text = MyData.dtmainEtibarbame.Rows[0]["c4"].ToString();

                }
                catch { }

                MyData.selectCommand("baza.accdb", "SELECT * FROM etibarnamearxiv WHERE a16 LIKE '%" + t2.Text + "%'");
                MyData.dtmainEtibarnameArxiv= new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainEtibarnameArxiv);
                dataGridView1.DataSource = MyData.dtmainEtibarnameArxiv;

                if (dataGridView1.Rows.Count > 0)
                {
                    dataGridView1.Rows[dataGridView1.Rows.Count - 1].Selected = true;
                    checkBox2.Checked = true;
                }
                else checkBox2.Checked = false;

                try
                {
                    lbSur.Text = "S/V"; lbSexsiyyet.Text = "Ş/V"; lbFiziki.Text = "Fiziki şəxs"; lbSur.Left = 84; lbSexsiyyet.Left = 84;

                    MyData.selectCommand("baza.accdb", "Select * from etibarnamesurucu where a1 Like '%" + txtlizinqalan.Text + "%'");
                    MyData.dtmainSurucu= new DataTable();
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
                        lbSur.Text = "VÖEN"; lbSexsiyyet.Text = "Direktor"; lbFiziki.Text = "Hüquqi şəxs"; lbSur.Left = 65; lbSexsiyyet.Left = 54;

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

                    MyData.selectCommand("baza.accdb", "Select * from etibarnamesurucu where a5 Like '%" + txtLayihe.Text + "%'");
                    MyData.dtmainSurucu= new DataTable();
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
                    txtlizinqalan.Text = txtlizinqalan.Text.Substring(0, 1).ToUpper(MyChange.DilDeyisme) + txtlizinqalan.Text.Substring(1, txtlizinqalan.Text.Length - 1).ToLower(MyChange.DilDeyisme);
                }
                catch { }



                try
                {
                    lbSur.Text = "S/V"; lbSexsiyyet.Text = "Ş/V"; lbFiziki.Text = "Fiziki şəxs"; lbSur.Left = 84; lbSexsiyyet.Left = 84;

                    
                    
                   
                    MyData.selectCommand("baza.accdb", "Select * from etibarnamesurucu where a1 Like '%" + txtlizinqalan.Text + "%'");
                    MyData.dtmainSurucu= new DataTable();
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
                        lbSur.Text = "VÖEN"; lbSexsiyyet.Text = "Direktor"; lbFiziki.Text = "Hüquqi şəxs"; lbSur.Left = 65; lbSexsiyyet.Left = 54;

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
                    
                    
                   
                    MyData.selectCommand("baza.accdb", "Select * from etibarnamesurucu where a1 Like " + "'%" + txtsurucu.Text + "%'");
                    MyData.dtmainSurucu= new DataTable();
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
                    else { txtunvan2.Text = ""; txtSurvesiqe2.Text = ""; txtsexsiyyet2.Text = ""; }

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
                
                
               
                MyData.selectCommand("baza.accdb", "Select * from etibarnamesurucu where a1 Like " + "'%" + txtsurucu.Text + "%'");
                MyData.dtmainSurucu= new DataTable();
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

        private void button3_Click(object sender, EventArgs e)
        {
            MyData.insertCommand("baza.accdb", "INSERT INTO MMX (a1,a2,a3,a4,a5,a6,a7,a8,a9,a10)values("

                                                                                                + $"'{ t1.Text}',"
                                                                                                + $"'{t2.Text}',"
                                                                                                + $"'{t3.Text}',"
                                                                                                + $"'{t4.Text}',"
                                                                                                + $"'{t5.Text}',"
                                                                                                + $"'{t6.Text}',"
                                                                                                + $"'{txtsurucu.Text}',"
                                                                                                + $"'{t8.Text}',"
                                                                                                + $"'{t9.Text}',"
                                                                                                + $"'{t10.Text}')");

            SendMail();

            MessageBox.Show("Yeni məlumat əlavə edildi");
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            
           
            MyData.selectCommand("baza.accdb", "Select * from Unvanlar where a1 Like '%" + comboBox2.Text + "%'");
            MyData.dtmainUnvanlar= new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainUnvanlar);
            try { txtunvan.Text = MyData.dtmainUnvanlar.Rows[0]["a1"].ToString(); }
            catch { txtunvan.Text = ""; }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                Directory.CreateDirectory("X:\\Umumi Senedler\\MMX\\" + t2.Text + "\\"); File.Copy("MMX\\POCT DYP.xlsx", $"X:\\Umumi Senedler\\MMX\\{t2.Text}\\POCT {t2.Text} ({t5.Text}).xlsx", true);
            }
            catch { }

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open($"X:\\Umumi Senedler\\MMX\\{t2.Text}\\POCT {t2.Text} ({t5.Text}).xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];

            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            oSheet.Cells[2, 3] = comboBox1.Text;
            oSheet.Cells[4, 3] = "İnzibati xəta törətmiş sürücülər barədə məlumatların verilməsi.";
            oSheet.Cells[5, 3] = txtunvan.Text;

            if (txtunvan.Text == txtunvan.Text) oSheet.Cells[5, 3] = $"{txtsurucu.Text}{Environment.NewLine}Ünvan: {txtunvan.Text}";

            oSheet.Cells[8, 3] = DateTime.Today.ToShortDateString();
            //  oSheet.PrintOut();
            //   oWB.Close(SaveChanges: false);
            // oXL.Application.Quit();

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
            //oSheet.PrintOut();
            //oWB.Close(SaveChanges: false);
            //oXL.Application.Quit();
        }

        private void MMXmelumat_Load(object sender, EventArgs e)
        {
            MMXMelumatNomre();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            MMXMelumatNomre();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (checkBox5.Checked == true) EtibarnameCap();
            else MelumatSurucu();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text.Substring(comboBox1.Text.Length - 3, 3) == "ASC") MektubAGBank();
            else MektubAGLizinq();

            MyData.updateCommand("baza.accdb", $"UPDATE MMXMelumatNomre  SET a1 ='{txtNomre.Text}'");
            MMXMelumatNomre();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text.Substring(comboBox1.Text.Length - 3, 3) == "ASC") BildirisAGBank();
            else BildirisAGLizinq();

            MyData.updateCommand("baza.accdb", $"UPDATE MMXMelumatNomre  SET a1 ='{txtNomre.Text}'");
            MMXMelumatNomre();
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

    }
}
