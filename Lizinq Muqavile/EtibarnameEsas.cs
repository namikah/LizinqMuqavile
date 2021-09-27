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
using System.IO;
using System.Net;
using System.Web;
using Nsoft;

namespace Lizinq_Muqavile
{
    public partial class EtibarnameEsas : Form
    {
        Excel.Application oXL;
        Excel._Workbook oWB;
        Excel._Worksheet oSheet;

        public EtibarnameEsas()
        {
            InitializeComponent();
        }

        private void myrefresh()
        {
            Cursor.Current = Cursors.WaitCursor;

            try
            {
                string commandText = "SELECT * FROM etibarnameneqliyyat WHERE  1=1";
                if (radioButton3.Checked == true)
                {
                    commandText += " and c3 like '%" + textBox1.Text + "%'";
                    commandText += " or c1 like '%" + textBox1.Text + "%'";
                    commandText += " or c13 like '%" + textBox1.Text + "%'";
                }

                if (radioButton4.Checked == true)
                {
                    commandText += " and c4 like '%" + textBox1.Text + "%'";
                    commandText += " or c2 like '%" + textBox1.Text + "%'";
                }
                //commandText += " order by c14 desc";

                MyData.selectCommand("baza.accdb", commandText);
                MyData.dtmain = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmain);
                dataGridView1.DataSource = MyData.dtmain;
            }
            catch { };
            Cursor.Current = Cursors.Default;

        }

        private void suruculer()
        {
            Cursor.Current = Cursors.WaitCursor;

            MyData.selectCommand("baza.accdb", "SELECT * FROM etibarnamesurucu");
            MyData.dtmainSuruculer = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainSuruculer);
            dataGridView2.DataSource = MyData.dtmainSuruculer;

            Cursor.Current = Cursors.Default;

        }

        private void arxivrefresh()
        {
            Cursor.Current = Cursors.WaitCursor;
            string commandText = "SELECT * FROM etibarnamearxiv WHERE 1=1";
            commandText += " and a16 like '%" + textBox2.Text + "%'";
            commandText += " or a8 like '%" + textBox2.Text + "%'";
            commandText += " or a2 like '%" + textBox2.Text + "%'";
            //commandText += " order by Kod desc";

            MyData.selectCommand("baza.accdb",commandText);
            MyData.dtmainArxiv = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainArxiv);
            dataGridView3.DataSource = MyData.dtmainArxiv;

            Cursor.Current = Cursors.Default;

        }

        private void etibarnamenomre()
        {
            Cursor.Current = Cursors.WaitCursor;

            MyData.selectCommand("baza.accdb", "Select * from etibarnamenomre");
            MyData.dtmainEtibarnameNomre = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainEtibarnameNomre);

            txtnomre.Text = (Convert.ToInt32(MyData.dtmainEtibarnameNomre.Rows[0][0]) + 1).ToString();

         try
            {
                if (Convert.ToInt32(MyData.dtmainEtibarnameNomre.Rows[0][0]) > 0 && Convert.ToInt32(MyData.dtmainEtibarnameNomre.Rows[0][0]) < 10) txtseriyaAGL.Text = "00000" + (Convert.ToInt32(MyData.dtmainEtibarnameNomre.Rows[0][0]) + 1).ToString();


                else if (Convert.ToInt32(MyData.dtmainEtibarnameNomre.Rows[0][0]) > 9 && Convert.ToInt32(MyData.dtmainEtibarnameNomre.Rows[0][0]) < 100) txtseriyaAGL.Text = "0000" + (Convert.ToInt32(MyData.dtmainEtibarnameNomre.Rows[0][0]) + 1).ToString();


                else if (Convert.ToInt32(MyData.dtmainEtibarnameNomre.Rows[0][0]) > 99 && Convert.ToInt32(MyData.dtmainEtibarnameNomre.Rows[0][0]) < 1000) txtseriyaAGL.Text = "000" + (Convert.ToInt32(MyData.dtmainEtibarnameNomre.Rows[0][0]) + 1).ToString();


                else if (Convert.ToInt32(MyData.dtmainEtibarnameNomre.Rows[0][0]) > 999 && Convert.ToInt32(MyData.dtmainEtibarnameNomre.Rows[0][0]) < 10000) txtseriyaAGL.Text = "00" + (Convert.ToInt32(MyData.dtmainEtibarnameNomre.Rows[0][0]) + 1).ToString();


                else if (Convert.ToInt32(MyData.dtmainEtibarnameNomre.Rows[0][0]) > 9999 && Convert.ToInt32(MyData.dtmainEtibarnameNomre.Rows[0][0]) < 100000) txtseriyaAGL.Text = "0" + (Convert.ToInt32(MyData.dtmainEtibarnameNomre.Rows[0][0]) + 1).ToString();


                else if (Convert.ToInt32(MyData.dtmainEtibarnameNomre.Rows[0][0]) > 99999 && Convert.ToInt32(MyData.dtmainEtibarnameNomre.Rows[0][0]) < 1000000) txtseriyaAGL.Text = (Convert.ToInt32(MyData.dtmainEtibarnameNomre.Rows[0][0]) + 1).ToString();
            }
            catch { };
            Cursor.Current = Cursors.Default;

        }

        private void VerilmisSonEtibarname()
        {
            Cursor.Current = Cursors.WaitCursor;

            btSon1.Text = "";
            btSon2.Text = "";
            btSon3.Text = "";

            //son 3 etibarname ucun
            MyData.selectCommand("baza.accdb", "SELECT * FROM etibarnamearxiv WHERE a16 LIKE '%" + txtnomresi.Text + "%'");
            MyData.dtmainArxiv = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainArxiv);
            dataGridView3.DataSource = MyData.dtmainArxiv;

            try { btSon1.Text = MyData.dtmainArxiv.Rows[MyData.dtmainArxiv.Rows.Count - 1]["a8"].ToString() + " (" + MyData.dtmainArxiv.Rows[MyData.dtmainArxiv.Rows.Count - 1]["a2"].ToString() + " - " + MyData.dtmainArxiv.Rows[MyData.dtmainArxiv.Rows.Count - 1]["a17"].ToString() + ")";}
            catch { };
            try { btSon2.Text = MyData.dtmainArxiv.Rows[MyData.dtmainArxiv.Rows.Count - 2]["a8"].ToString() + " (" + MyData.dtmainArxiv.Rows[MyData.dtmainArxiv.Rows.Count - 2]["a2"].ToString() + " - " + MyData.dtmainArxiv.Rows[MyData.dtmainArxiv.Rows.Count - 2]["a17"].ToString() + ")"; }
            catch { };
            try { btSon3.Text = MyData.dtmainArxiv.Rows[MyData.dtmainArxiv.Rows.Count - 3]["a8"].ToString() + " (" + MyData.dtmainArxiv.Rows[MyData.dtmainArxiv.Rows.Count - 3]["a2"].ToString() + " - " + MyData.dtmainArxiv.Rows[MyData.dtmainArxiv.Rows.Count - 3]["a17"].ToString() + ")"; }
            catch { };

            Cursor.Current = Cursors.Default;

        }

        private void DYPUmumiMelumat()
        {
            Cursor.Current = Cursors.WaitCursor;

            MyData.selectCommand("baza.accdb", "Select * From DYPUmumiMelumat");
            MyData.dtmainDYPUmumiMelumat = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainDYPUmumiMelumat);

            try
            {
                File.Copy("MMX\\DYP Umumi Melumat.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\DYP Umumi Melumat.xlsx", true);
            }
            catch { MessageBox.Show("DYP Umumi Melumat.xlsx tapılmadı."); }

            int a, b;
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\DYP Umumi Melumat.xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;

            oSheet.Cells[1, 1] = "№";
            oSheet.Cells[1, 2] = "Dövlət Nömrə Nişanı";
            oSheet.Cells[1, 3] = "Marka";
            oSheet.Cells[1, 4] = "Lizinq alan";
            oSheet.Cells[1, 5] = "Sürücülər və məsul şəxslər.";
            oSheet.Cells[1, 6] = "Telefon";

            for (a = 0; a < MyData.dtmainDYPUmumiMelumat.Rows.Count; a++)
            {
                oSheet.Range["A" + (a+2)].Select();

                for (b = 0; b < 6; b++)
                {
                    oSheet.Cells[a + 2, b + 1] = MyData.dtmainDYPUmumiMelumat.Rows[a][b].ToString();
                   
                }

                oSheet.Range["A" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["B" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["C" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["D" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["E" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["F" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;

            }

            oSheet.Range["A" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["B" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["C" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["D" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["E" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["F" + 1].Borders.LineStyle = Excel.Constants.xlSolid;

            Cursor.Current = Cursors.Default;

        }

        public void WordDoc()
        {
            Cursor.Current = Cursors.WaitCursor;

            try { File.Copy("etibarname.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Etibarname - " + txtnomresi.Text + " " + txtSAA.Text + ".doc", true); }
            catch { MessageBox.Show("'\\192.168.10.5\\Common\\AGLizinq\\etibarname.doc' tapılmadı."); }

            object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\Etibarname - " + txtnomresi.Text + " " + txtSAA.Text + ".doc";


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


                if (radioButton1.Checked == true)
                {
                MyChange.FindAndReplace(word, "000", "№ 000/" + txtTarix.Text.Substring(txtTarix.Text.Length-2,2) + Environment.NewLine + Environment.NewLine + txtTarix.Text + "-cı il");
                MyChange.FindAndReplace(word, "111", " “AGLizinq” QAPALI SƏHMDAR CƏMİYYƏTİ");
                MyChange.FindAndReplace(word, "222", " “AGLizinq” QSC");
                MyChange.FindAndReplace(word, "5555", " “AGLizinq” QSC-nin" + Environment.NewLine + "Baş meneceri vəzifəsini icra edən                    				R.İ.Musayev");
                }

                if (radioButton2.Checked == true)
                {
                MyChange.FindAndReplace(word, "000", "");
                MyChange.FindAndReplace(word, "111", " “AGBank” AÇIQ SƏHMDAR CƏMİYYƏTİ");
                MyChange.FindAndReplace(word, "222", " “AGBank” ASC");
                MyChange.FindAndReplace(word, "5555", " “AGBank” ASC-nin" + Environment.NewLine + "Idarə Heyətinin Sədri" + Environment.NewLine + "Ə.S.Cəlilov" + Environment.NewLine + "(17.08.17 tarixli etibarnaməyə əsasən" + Environment.NewLine + "“AGLizinq” QSC-nin Baş meneceri" + Environment.NewLine + "vəzifəsini icra edən                     				        		R.İ Musayev");
                }

                MyChange.FindAndReplace(word, "111111111", txtUnvanimiz.Text);
                MyChange.FindAndReplace(word, "333", txttexpassnomre.Text);
                MyChange.FindAndReplace(word, "444", txtmarka.Text);

                if (lbban.Text == "Ban №-si") { MyChange.FindAndReplace(word, "555", txtBan.Text); MyChange.FindAndReplace(word, "55", "BAN"); }
                if (lbban.Text == "Şassi N-si" || lbban.Text == "Şassi №-si") { MyChange.FindAndReplace(word, "555", txtBan.Text); MyChange.FindAndReplace(word, "55", "ŞASSİ"); }
                if (lbsassi.Text == "Ban №-si") { MyChange.FindAndReplace(word, "555", txtSassi.Text); MyChange.FindAndReplace(word, "55", "BAN"); }
                if (lbsassi.Text == "Şassi N-si" || lbsassi.Text == "Şassi №-si") { MyChange.FindAndReplace(word, "555", txtSassi.Text); MyChange.FindAndReplace(word, "55", "ŞASSİ"); }

                MyChange.FindAndReplace(word, "666", txtBuraxilis.Text);
                MyChange.FindAndReplace(word, "777", dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c5"].Value.ToString());
                MyChange.FindAndReplace(word, "888", txtnomresi.Text);
                MyChange.FindAndReplace(word, "999", txtSAA.Text + "na");
                MyChange.FindAndReplace(word, "999", txtSAA.Text + "na");
                MyChange.FindAndReplace(word, "1111", txtSuruculuk.Text);
                MyChange.FindAndReplace(word, "2222", txtUnvan.Text);
                MyChange.FindAndReplace(word, "3333", txtTarix.Text);
                MyChange.FindAndReplace(word, "4444", txtEtibarnameBitme.Text);

                Cursor.Current = Cursors.Default;

        }

        public void WordDocHuquqi()
        {
            Cursor.Current = Cursors.WaitCursor;


            try { File.Copy("Etibarname Huquqi Sexs.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Etibarname Huquqi - " + txtnomresi.Text + ".doc", true); }
            catch { MessageBox.Show("'\\192.168.10.5\\Common\\AGLizinq\\etibarname.doc' tapılmadı."); }

            object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\Etibarname Huquqi - " + txtnomresi.Text + ".doc";


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

            string c = MyChange.ReqemToMetn(Convert.ToInt32(cbMuddet.Text));

            MyChange.FindAndReplace(word, "000", "TARİX: " + txtTarix.Text + "-cı il");
            MyChange.FindAndReplace(word, "222", txtTexpasstarix.Text + "-ci il");
            MyChange.FindAndReplace(word, "222", txtTexpasstarix.Text + "-ci ildə");
            MyChange.FindAndReplace(word, "333", txtEsas.Text);
            MyChange.FindAndReplace(word, "444", dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c3"].Value.ToString());
            MyChange.FindAndReplace(word, "555", txtSuruculuk.Text);
            MyChange.FindAndReplace(word, "666", txtUnvan.Text);
            MyChange.FindAndReplace(word, "777", txttexpassnomre.Text);
          
            if (radioButton1.Checked == true)
            {
                MyChange.FindAndReplace(word, "111", "“AGLizinq” QSC " + "(VÖEN: 1300616961, Ünvan: " + txtUnvanimiz.Text + ")");
                MyChange.FindAndReplace(word, "888", "“AGLizinq” QSC");
            }

            if (radioButton2.Checked == true)
            {
                MyChange.FindAndReplace(word, "111", "“AGBank” ASC " + "(VÖEN: 9900019651, Ünvan: " + txtUnvanimiz.Text + ")");
                MyChange.FindAndReplace(word, "888", "“AGBank” ASC");
            }

            MyChange.FindAndReplace(word, "999", txtmarka.Text);
            MyChange.FindAndReplace(word, "1000", txtBan.Text);
            MyChange.FindAndReplace(word, "1001", txtSassi.Text);
            MyChange.FindAndReplace(word, "2000", dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c5"].Value.ToString());
            MyChange.FindAndReplace(word, "3000", txtBuraxilis.Text);
            MyChange.FindAndReplace(word, "4000", txtnomresi.Text);
            MyChange.FindAndReplace(word, "5000", cbMuddet.Text + " (" + c + ") " + cbAyGun.Text);
            MyChange.FindAndReplace(word, "6000", txtEtibarnameBitme.Text + "-ci ilədək");

            Cursor.Current = Cursors.Default;

        }

        private void EtibarnameEsas_Load(object sender, EventArgs e)
        {
            MyChange.SetKeyboardLayout(MyChange.GetInputLanguageByName("AZ"));

            dtEtibarnameBugun.Value = DateTime.Now;
            base.Text = Environment.UserName + "  /  Etibarnamə";

            myrefresh();
            suruculer();
            arxivrefresh();
            etibarnamenomre();

            lbxeberler.Left = base.Width;
            lbxeberler.Text = dataGridView1.Rows.Count.ToString() + " nəqliyyat vasitəsi və " + dataGridView2.Rows.Count.ToString() + " sürücü qeydiyyata alınıb.";
            txtarxivmelumat.Text = dataGridView1.Rows.Count.ToString() + " nəqliyyat vasitəsi və " + dataGridView2.Rows.Count.ToString() + " sürücü qeydiyyata alınmışdır.";
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            txtBiz.Text = "“AGLizinq” QSC  VÖEN: 1300616961";
            txtUnvanimiz.Text = "AZ1073 Bakı şəh, Yasamal r-nu, Landau küç 16";
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            txtBiz.Text = "“AGBank” ASC  VÖEN: 9900019651";
            txtUnvanimiz.Text = "AZ1022 Bakı şəh, Nəsimi r-nu," + Environment.NewLine + "Cəlil Məmmədquluzadə, ev 102A";
        }

        private void button2_Click(object sender, EventArgs e)
        {
           

            DialogResult result2 = MessageBox.Show("Davam etmək istəyirsiniz?", "?", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (result2 == DialogResult.No) { return; } 
            Cursor.Current = Cursors.WaitCursor;

            try
            {
                File.Copy("Etibarname.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Etibarname - " + txtnomresi.Text.Substring(0, 9).ToString() + ".xlsx", true);
            }
            catch { MessageBox.Show("Etibarname.xlsx tapılmadı."); }

            DateTime dt = DateTime.Now;
            int azn = 0;
            string c = "";

            try
            {
                if (cbAyGun.Text == "ay")
                {
                    azn = Convert.ToInt32(cbMuddet.Text) * 3;
                }
            }
            catch { }
            try
            {
                c = MyChange.ReqemToMetn(Convert.ToInt32(cbMuddet.Text));
            }
            catch { };

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Etibarname - " + txtnomresi.Text.Substring(0, 9).ToString() + ".xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

            oSheet.Cells[1, 4] = txtseriyaAGL.Text;
            oSheet.Cells[3, 3] = "'" + txtTarix.Text;
            oSheet.Cells[5, 2] = txtBiz.Text;
            if (radioButton1.Checked == true) oSheet.Cells[3, 7] = "Şəhadətnaməsi əsasında " + radioButton1.Text + "-yə məxsus";
            else if (radioButton2.Checked == true) oSheet.Cells[3, 7] = "Şəhadətnaməsi əsasında " + radioButton2.Text + "-yə məxsus";

            if (txtUnvanimiz.Text.Length > 45)
            {
                oSheet.Cells[7, 2] = txtUnvanimiz.Text.Substring(0, 29);
                oSheet.Cells[8, 2] = txtUnvanimiz.Text.Substring(29, txtUnvanimiz.Text.Length - 29);
            }
            else
            {
                oSheet.Cells[7, 2] = txtUnvanimiz.Text;
            }

            oSheet.Cells[9, 2] = txtEsas.Text;

            if (txtSuruculuk.Text != "")
            {
                oSheet.Cells[12, 1] = "Sürücülük vəsiqəsi:";
                oSheet.Cells[12, 4] = txtSuruculuk.Text;
            }
            else if (txtSexsiyyet.Text != "")
            {
                oSheet.Cells[12, 1] = "Şəxsiyyət vəsiqəsi:";
                oSheet.Cells[12, 4] = txtSexsiyyet.Text;
            }

            oSheet.Cells[14, 1] = txtUnvan.Text;
            try
            {
                if (txtUnvan.Text.Length > 50)
                {
                    for (int y = 30; y < txtUnvan.Text.Length; y++)
                    {
                        if (txtUnvan.Text.Substring(y, 1) == ",")
                        {
                            oSheet.Cells[14, 1] = txtUnvan.Text.Substring(0, y + 1);
                            oSheet.Cells[16, 1] = txtUnvan.Text.Substring(y + 2, txtUnvan.Text.Length - y - 2);
                            y = txtUnvan.Text.Length;
                        }

                    }
                }
            }
            catch { }
            oSheet.Cells[18, 1] = txtSAA.Text;
            oSheet.Cells[23, 1] = "'" + txtTexpasstarix.Text;
            oSheet.Cells[1, 7] = txttexpassnomre.Text;
            oSheet.Cells[5, 7] = txtmarka.Text;
            oSheet.Cells[7, 10] = txtmuherrik.Text;
            if (txtBan.Text != "") { oSheet.Cells[8, 8] = txtBan.Text; oSheet.Cells[8, 7] = lbban.Text; }
            if (txtSassi.Text != "") { oSheet.Cells[8, 8] = txtSassi.Text; oSheet.Cells[8, 7] = lbsassi.Text; }

            oSheet.Cells[9, 8] = txtBuraxilis.Text;
            oSheet.Cells[11, 8] = txtnomresi.Text;
            oSheet.Cells[20, 11] = "'" + txtEtibarnameBitme.Text;
            if (c != "") oSheet.Cells[19, 9] = cbMuddet.Text + " (" + c + ") " + cbAyGun.Text;
            else oSheet.Cells[19, 9] = cbMuddet.Text + " " + c + " " + cbAyGun.Text;

            oSheet.Cells[20, 14] = "  - ci";

            try
            {

                if (cbCOPY.Checked == true)
                {
                    MyData.selectCommand("baza.accdb", "Select * From etibarnamearxiv where a1=" + "'" + txtseriyaAGL.Text + "'");
                    MyData.dtmainArxiv = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmainArxiv);

                    if (MyData.dtmainArxiv.Rows.Count > 0)
                    {
                        MyData.updateCommand("baza.accdb", "UPDATE etibarnamearxiv SET "
                                                                                                            + "a1 = '" + txtseriyaAGL.Text + "',"
                                                                                                            + "a2 = '" + txtTarix.Text + "',"
                                                                                                            + "a3 = '" + txtBiz.Text.Substring(0, 14).ToString() + "',"
                                                                                                            + "a4 = '" + txtEsas.Text + "',"
                                                                                                            + "a5 = '" + txtSexsiyyet.Text + "',"
                                                                                                            + "a6 = '" + txtSuruculuk.Text + "',"
                                                                                                            + "a7 = '" + txtUnvan.Text + "',"
                                                                                                            + "a8 = '" + txtSAA.Text + "',"
                                                                                                            + "a9 = '" + txtTexpasstarix.Text + "',"
                                                                                                            + "a10 = '" + txttexpassnomre.Text + "',"
                                                                                                            + "a11 = '" + txtmarka.Text + "',"
                                                                                                            + "a12 = '" + txtmuherrik.Text + "',"
                                                                                                            + "a13 = '" + txtBan.Text + "',"
                                                                                                            + "a14 = '" + txtSassi.Text + "',"
                                                                                                            + "a15 = '" + txtBuraxilis.Text + "',"
                                                                                                            + "a16 = '" + txtnomresi.Text + "',"
                                                                                                            + "a17 = '" + txtEtibarnameBitme.Text + "',"
                                                                                                            + "a18 = '" + azn + "' where a1=" + "'" + txtseriyaAGL.Text + "'");

                        arxivrefresh();
                    }
                    else
                    {
                        MyData.insertCommand("baza.accdb", "INSERT INTO etibarnamearxiv (a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13,a14,a15,a16,a17,a18)values("


                                                                                                            + "'" + txtseriyaAGL.Text + "',"
                                                                                                            + "'" + txtTarix.Text + "',"
                                                                                                            + "'" + txtBiz.Text.Substring(0, 14).ToString() + "',"
                                                                                                            + "'" + txtEsas.Text + "',"
                                                                                                            + "'" + txtSexsiyyet.Text + "',"
                                                                                                            + "'" + txtSuruculuk.Text + "',"
                                                                                                            + "'" + txtUnvan.Text + "',"
                                                                                                            + "'" + txtSAA.Text + "',"
                                                                                                            + "'" + txtTexpasstarix.Text + "',"
                                                                                                            + "'" + txttexpassnomre.Text + "',"
                                                                                                            + "'" + txtmarka.Text + "',"
                                                                                                            + "'" + txtmuherrik.Text + "',"
                                                                                                            + "'" + txtBan.Text + "',"
                                                                                                            + "'" + txtSassi.Text + "',"
                                                                                                            + "'" + txtBuraxilis.Text + "',"
                                                                                                            + "'" + txtnomresi.Text + "',"
                                                                                                            + "'" + txtEtibarnameBitme.Text + "',"
                                                                                                            + "'" + azn + "')");
                        arxivrefresh();

                    }

                    MyData.updateCommand("baza.accdb", "UPDATE etibarnamenomre  SET a1 =" + "'" + txtseriyaAGL.Text + "'");
                    etibarnamenomre();

                    lbetibarnamemeblegi.Visible = false;
                    lbetibarnamemeblegi.Text = "";
                }
            }
            catch { MessageBox.Show("Kopyalama alinmadı"); };

            //emeliyyatlar ucun
            MyData.insertCommand("baza.accdb", "insert into Emeliyyatlar (a1, a2, a3) Values ('" + DateTime.Now + "', 'ETIBARNAMƏ - " + txtnomresi.Text + " - " + txtSAA.Text + " - " + cbMuddet.Text + " " + cbAyGun.Text + "','" + Environment.MachineName + "')");

            VerilmisSonEtibarname();

            Cursor.Current = Cursors.Default;
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Cursor.Current = Cursors.WaitCursor;
                progressBar2.Value = 0;
                progressBar2.Maximum = 100;
                progressBar2.Step = 50;

                try
                {
                    textBox1.Text = textBox1.Text.Substring(0, 1).ToUpper(MyChange.DilDeyisme) + textBox1.Text.Substring(1, textBox1.Text.Length - 1).ToLower(MyChange.DilDeyisme);
                }
                catch { }


                myrefresh(); progressBar2.PerformStep();
                label2.ForeColor = Color.Black;
                VerilmisSonEtibarname(); progressBar2.PerformStep();
                Cursor.Current = Cursors.Default;
            }
        }
        
        private void nəqliyyatVasitələriToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                File.Copy("Bos.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Neqliyyat - " + ".xlsx", true);
            }
            catch { MessageBox.Show("Bos.xlsx tapılmadı."); }

            int a, b;

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Neqliyyat - " + ".xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
           
            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;


            oSheet.Cells[1, 1] = "Dövlət Nömrə Nişanı";
            oSheet.Cells[1, 2] = "Ban №-si";
            oSheet.Cells[1, 3] = "Mühərrik №-si";
            oSheet.Cells[1, 4] = "Şassi №-si";
            oSheet.Cells[1, 5] = "Tex. Pass. №-si";
            oSheet.Cells[1, 6] = "Buraxılış ili";
            oSheet.Cells[1, 7] = "Tex. Pass. Verilmə tarixi";
            oSheet.Cells[1, 8] = "Marka";
            oSheet.Cells[1, 9] = "Lizinq alan";
            oSheet.Cells[1, 10] = "Layihə nömrəsi";
            oSheet.Cells[1, 11] = "Zavod №- si";

            for (a = 0; a < MyData.dtmain.Rows.Count; a++)
            {

                for (b = 1; b < 12; b++)
                {
                        oSheet.Cells[a+2, b] = MyData.dtmain.Rows[a][b].ToString();

                }
                oSheet.Range["A" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["B" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["C" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["D" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["E" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["F" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["G" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["H" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["I" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["J" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["K" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
            }

            oSheet.Range["A" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["B" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["C" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["D" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["E" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["F" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["G" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["H" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["I" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["J" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["K" + 1].Borders.LineStyle = Excel.Constants.xlSolid;

         //   oSheet.PrintOut();
          //  oWB.Close(SaveChanges: false);
          //  oXL.Workbooks.Close();
           // oXL.Application.Quit();
          //  oXL.Quit();
            oSheet.Columns.AutoFit();
            oSheet.Rows.AutoFit();
        }

        private void sürücülərToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                File.Copy("Bos.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Suruculer - " + ".xlsx", true);
            }
            catch { MessageBox.Show("Bos.xlsx tapılmadı."); }

            suruculer();
            int a, b;

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Suruculer - " + ".xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;

            oSheet.Cells[1, 1] = "№";
            oSheet.Cells[1, 2] = "Sürücü";
            oSheet.Cells[1, 3] = "Şəxsiyyət vəsiqəsi:";
            oSheet.Cells[1, 4] = "Sürücülük vəsiqəsi:";
            oSheet.Cells[1, 5] = "Ünvan";
            oSheet.Cells[1, 6] = "Layihə №";
            oSheet.Cells[1, 7] = "Vəsiqənin bitmə tarixi";
            oSheet.Cells[1, 8] = "E-mail";
            oSheet.Cells[1, 9] = "Telefon";
            oSheet.Cells[1, 10] = "Mobile";

            for (a = 0; a < MyData.dtmainSuruculer.Rows.Count; a++)
            {
                for (b = 0; b < 10; b++)
                {
                    oSheet.Cells[a + 2, b + 1] = dataGridView2.Rows[a].Cells[b].Value.ToString();

                }

                oSheet.Range["A" + (a  + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["B" + (a  + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["C" + (a  + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["D" + (a  + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["E" + (a  + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["F" + (a  + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["G" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["H" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["I" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["J" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
               
            }

            
            oSheet.Range["A" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["B" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["C" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["D" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["E" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["F" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["G" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["H" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["I" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["J" + 1].Borders.LineStyle = Excel.Constants.xlSolid;


         //   oSheet.PrintOut();
       //     oWB.Close(SaveChanges: false);
        //    oXL.Workbooks.Close();
         //   oXL.Application.Quit();
       //     oXL.Quit()

            oSheet.Columns.AutoFit();
            oSheet.Rows.AutoFit();

            myrefresh();
        }

        private void çıxışToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult result2 = MessageBox.Show("Davam etmək istəyirsiniz?", "?", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (result2 == DialogResult.No) { return; }

            base.Close();
        }

        private void helpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Press F1");
        }

        private void dataGridView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            timer1.Enabled = true;
            // CreateSqlConnection();

            string commandText = "SELECT * FROM etibarnamesurucu WHERE 1=1";
            try
            {
                commandText += " and a5 like '%" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c4"].Value.ToString() + "%'";
            
            }   catch { };
            MyData.selectCommand("baza.accdb", commandText);
            MyData.dtmainSuruculer = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainSuruculer);
            dataGridView2.DataSource = MyData.dtmainSuruculer;

            VerilmisSonEtibarname();
        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            btvesiqe.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);

            try
            {
                txtSAA.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["a1"].Value.ToString();
                txtSuruculuk.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["a3"].Value.ToString();
                txtSexsiyyet.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["a2"].Value.ToString();
                txtUnvan.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["a4"].Value.ToString();
                if (dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["a6"].Value.ToString() == "") btvesiqe.Text = "* Vəsiqənin bitmə tarixi (qeyd olunmayıb !!!)";
                else btvesiqe.Text = "* Vəsiqənin bitmə tarixi (" + dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["a6"].Value.ToString() + ")";
            }
            catch
            {
                txtSAA.Text = "";
                txtSuruculuk.Text = "";
                txtSexsiyyet.Text = "";
                txtUnvan.Text = "";
                btvesiqe.Text = "* Vəsiqənin bitmə tarixi";
            };

            try //vesiqenin bitme vaxtinin yoxlanmasi
            {
                DateTime bitme = Convert.ToDateTime(dtEtibarnameBitme.Value).Date;
                DateTime VesiqeBitme = Convert.ToDateTime(dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["a6"].Value).Date;

                if (bitme <= VesiqeBitme) { btvesiqe.ForeColor = Color.Green; }
                else { btvesiqe.ForeColor = Color.Red; return; }
            }
            catch { }
        }

        private void textBox14_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    textBox14.Text = textBox14.Text.Substring(0, 1).ToUpper(MyChange.DilDeyisme) + textBox14.Text.Substring(1, textBox14.Text.Length - 1).ToLower(MyChange.DilDeyisme);
                }
                catch { }

                try
                {
                    MyData.selectCommand("baza.accdb", "SELECT * FROM etibarnamesurucu WHERE a5 like '%" + textBox14.Text + "%' or a1 like '%" + textBox14.Text + "%'");
                    MyData.dtmainSuruculer = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmainSuruculer);
                    dataGridView2.DataSource = MyData.dtmainSuruculer;
                }
                catch { };

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult result2 = MessageBox.Show("Davam etmək istəyirsiniz?", "?", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (result2 == DialogResult.No) { return; }

            string s1, s2, s3, s4, s5, s6, s7, s8, s9, s10, s11, s12, s13;
            



                                                                                                s1= "'" + txtnomre1.Text + "'";
                                                                                                s2= "'" + txtbannomre2.Text + "'";
                                                                                                s3= "'" + txtmuherriknomre3.Text + "'";
                                                                                                s4= "'" + txtsassinomre4.Text + "'";
                                                                                                s5= "'" + txttexpass5.Text + "'";
                                                                                                s6= "'" + txtburaxilisili6.Text + "'";
                                                                                                s7= "'" + txtpassverilmetarix7.Text + "'";
                                                                                                s8= "'" + txtmarkasi8.Text + "'";
                                                                                                s9= "'" + txtlizinqalanadi9.Text + "'";
                                                                                                s10= "'" + txtlayihensi10.Text + "'";
                                                                                                s11= "'" + txtzavodnomresi11.Text + "'";
                                                                                                s12= "'" + txtreng12.Text + "'";
                                                                                                s13 = "'" + txtQeyd.Text + "'";

            MyData.selectCommand("baza.accdb", "Select * from etibarnameneqliyyat where c8=" + s2 + "and c9=" + s3 + "and c10=" + s4);
            MyData.dtmain = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmain);

                if (MyData.dtmain.Rows.Count > 0)
                {
                MyData.updateCommand("baza.accdb","UPDATE etibarnameneqliyyat SET "
                                                                                         + "c1 =" + s1 + ","
                                                                                         + "c8 =" + s2 + ","
                                                                                         + "c9 =" + s3 + ","
                                                                                         + "c10 =" + s4 + ","
                                                                                         + "c12 =" + s5 + ","
                                                                                         + "c6 =" + s6 + ","
                                                                                         + "c7 =" + s7 + ","
                                                                                         + "c2 =" + s8 + ","
                                                                                         + "c3 =" + s9 + ","
                                                                                         + "c4 =" + s10 + ","
                                                                                         + "c11 =" + s11 + ","
                                                                                         + "c5 =" + s12 + ","
                                                                                         + "c13 =" + s13
                                                                                         + " WHERE c8=" + s2 + "and c9=" + s3 + "and c10=" + s4);

                    
                    MessageBox.Show("Məlumat yeniləndi");
                    myrefresh();
                    suruculer();
                    arxivrefresh();
                    return;


                }

            MyData.insertCommand("baza.accdb", "insert into etibarnameneqliyyat (c1,c8,c9,c10,c12,c6,c7,c2,c3,c4,c11,c5,c13)values("

                                                                                                + "'" + txtnomre1.Text + "',"
                                                                                                + "'" + txtbannomre2.Text + "',"
                                                                                                + "'" + txtmuherriknomre3.Text + "',"
                                                                                                + "'" + txtsassinomre4.Text + "',"
                                                                                                + "'" + txttexpass5.Text + "',"
                                                                                                + "'" + txtburaxilisili6.Text + "',"
                                                                                                + "'" + txtpassverilmetarix7.Text + "',"
                                                                                                + "'" + txtmarkasi8.Text + "',"
                                                                                                + "'" + txtlizinqalanadi9.Text + "',"
                                                                                                + "'" + txtlayihensi10.Text + "',"
                                                                                                + "'" + txtzavodnomresi11.Text + "',"
                                                                                                + "'" + txtreng12.Text + "',"
                                                                                                + "'" + txtQeyd.Text + "')");

            MessageBox.Show("Yeni məlumat əlavə edildi");
            myrefresh();
            suruculer();
            arxivrefresh();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DialogResult result2 = MessageBox.Show("Davam etmək istəyirsiniz?", "?", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (result2 == DialogResult.No) { return; }

            txtnomre1.Text = "";
            txtbannomre2.Text = "";
            txtmuherriknomre3.Text = "";
            txtsassinomre4.Text = "";
            txttexpass5.Text = "";
            txtburaxilisili6.Text = "";
            txtpassverilmetarix7.Text = "";
            txtmarkasi8.Text = "";
            txtlizinqalanadi9.Text = "";
            txtlayihensi10.Text = "";
            txtzavodnomresi11.Text = "";
            txtreng12.Text = "";
            txtQeyd.Text = "";
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DialogResult result2 = MessageBox.Show("Davam etmək istəyirsiniz?", "?", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (result2 == DialogResult.No) { return; }

            string s1, s2 = "''", s3 = "''", s4, s5, s6, s7, s8, s9;
            s1 = "'" + txtad.Text + "'";
            if (txtsexsiyyetseriya.Text != "" || txtsexsiyyetnomre.Text!="") s2 = "'" + txtsexsiyyetseriya.Text + " № " + txtsexsiyyetnomre.Text + "'";
            if (txtsurucuseriya.Text != "" || txtsurucunomre.Text != "") s3 = "'" + txtsurucuseriya.Text + " № " + txtsurucunomre.Text + "'";
            s4 = "'" + txtunvansurucu.Text + "'";
            s5 = "'" + txtlayihe.Text + "'";
            s6 = "'" + dttarix.Text + "'";
            s7 = "'" + txtemail.Text + "'";
            s8 = "'" + tel1.Text + "'";
            s9 = "'" + tel2.Text + "'";

            MyData.selectCommand("baza.accdb", "Select * from etibarnamesurucu where a1=" + s1);
            MyData.dtmainSuruculer = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainSuruculer);

            if (MyData.dtmainSuruculer.Rows.Count > 0)
            {
                string k = txtad.Text;
                MyData.updateCommand("baza.accdb", "UPDATE etibarnamesurucu SET "
                                                                                     + "a1 =" + s1 + ","
                                                                                     + "a2 =" + s2 + ","
                                                                                     + "a3 =" + s3 + ","
                                                                                     + "a4 =" + s4 + ","
                                                                                     + "a5 =" + s5 + ","
                                                                                     + "a6 =" + s6 + ","
                                                                                     + "a7 =" + s7 + ","
                                                                                     + "a8 =" + s8 + ","
                                                                                     + "a9 =" + s9
                                                                                     + " WHERE a1=" + s1);

                MyData.updateCommand("baza.accdb", "UPDATE Telefon SET "
                                                                                     + "c1 ='" + txtad.Text + "',"
                                                                                     + "c2 ='" + tel1.Text + "',"
                                                                                     + "c3 ='" + tel2.Text + "'"
                                                                                     + " WHERE c1='" + k + "'");

                MessageBox.Show("Məlumat yeniləndi");
                myrefresh();
                suruculer();
                arxivrefresh();
                return;


            }
            MyData.insertCommand("baza.accdb", "insert into etibarnamesurucu (a1,a2,a3,a4,a5,a6,a7,a8,a9)values("

                                                                                                + s1 + ","
                                                                                                + s2 + ","
                                                                                                + s3 + ","
                                                                                                + s4 + ","
                                                                                                + s5 + ","
                                                                                                + s6 + ","
                                                                                                + s7 + ","
                                                                                                + s8 + ","
                                                                                                + s9 + ")");
            suruculer();

            MyData.insertCommand("baza.accdb", "insert into Telefon (c1,c2,c3)values("

                                                                                                + "'" + txtad.Text + "',"
                                                                                                + "'" + tel1.Text + "',"
                                                                                                + "'" + tel2.Text + "')");

            MessageBox.Show("Yeni məlumat əlavə edildi");
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            DialogResult result2 = MessageBox.Show("Davam etmək istəyirsiniz?", "?", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (result2 == DialogResult.No) { return; }

            txtad.Text = "";
            txtsurucuseriya.Text = "";
            txtsurucunomre.Text = "";
            txtsexsiyyetseriya.Text = "";
            txtsexsiyyetnomre.Text = "";
            txtunvansurucu.Text = "";
            txtlayihe.Text = "";
            tel1.Text = "";
            tel2.Text = "";
            dttarix.Text = "01-01-2001";
            txtemail.Text = "";
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            lbban.Text = "Ban №-si";
            lbsassi.Text = "Şassi №-si";

            try
            {
                txtnomresi.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c1"].Value.ToString();

                //AGBANK yaxud AGLIZINQ layihesi oldugunu secmek ucun
                if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c13"].Value.ToString() == "AGL") radioButton1.Checked = true;
                else radioButton2.Checked = true;
            }
            catch { }
                //textBox1.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c1"].Value.ToString();

            try
            {
                txtBan.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c8"].Value.ToString();
                txtmuherrik.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c9"].Value.ToString();
                txtSassi.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c10"].Value.ToString();
                txttexpassnomre.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c12"].Value.ToString();
                txtBuraxilis.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c6"].Value.ToString();
                txtTexpasstarix.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c7"].Value.ToString();
                txtmarka.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c2"].Value.ToString();
                txtEsas.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c4"].Value.ToString();
            }
            catch { }

            try
            {
                if (txtBan.Text == "") { txtBan.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c11"].Value.ToString(); lbban.Text = "Zavod №-si"; return; }
                if (txtSassi.Text == "") { txtSassi.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c11"].Value.ToString(); lbsassi.Text = "Zavod №-si"; }
            }
            catch { }

                VerilmisSonEtibarname();

        }

        private void tabControl1_Selected(object sender, TabControlEventArgs e)
        {
            if (dataGridView1.Height >= base.ClientSize.Height - 168) { timer1.Enabled = true; }
            if (dataGridView1.Height < base.ClientSize.Height - 170) { timer2.Enabled = true; }

            timer2.Enabled = true;
            if (tabControl1.SelectedTab == tabPage2) { timer2.Enabled = false; timer1.Enabled = true; }
        }

        private void editToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try { txtnomre1.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c1"].Value.ToString(); } 
            catch { };
            try { txtbannomre2.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c8"].Value.ToString(); }
            catch { };
            try { txtmuherriknomre3.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c9"].Value.ToString(); }
            catch { };
            try { txtsassinomre4.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c10"].Value.ToString(); }
            catch { };
            try { txttexpass5.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c12"].Value.ToString(); }
            catch { };
            try { txtburaxilisili6.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c6"].Value.ToString(); }
            catch { };
            try { txtpassverilmetarix7.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c7"].Value.ToString(); }
            catch { };
            try { txtmarkasi8.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c2"].Value.ToString(); }
            catch { };
            try { txtlizinqalanadi9.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c3"].Value.ToString(); }
            catch { };
            try { txtlayihensi10.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c4"].Value.ToString(); }
            catch { };
            try { txtzavodnomresi11.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c11"].Value.ToString(); }
            catch { };
            try { txtreng12.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c5"].Value.ToString(); }
            catch { };
            try { txtQeyd.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c13"].Value.ToString(); }
            catch { };


          tabControl1.SelectedTab = tabPage1;

        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try { txtad.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["a1"].Value.ToString(); }
            catch { }
            try { txtsexsiyyetseriya.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["a2"].Value.ToString().Substring(0, 3); } 
            catch { };
            try { txtsexsiyyetnomre.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["a2"].Value.ToString().Substring(6, dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[2].Value.ToString().Length - 6); }
            catch { };
            try { txtsurucuseriya.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["a3"].Value.ToString().Substring(0, 2); }
            catch { };
            try { txtsurucunomre.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["a3"].Value.ToString().Substring(5, dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[3].Value.ToString().Length - 5); }
            catch { };
            try { txtunvansurucu.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["a4"].Value.ToString(); }
            catch { }
            try { txtlayihe.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["a5"].Value.ToString(); }
            catch { }
            try { dttarix.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["a6"].Value.ToString(); }
            catch { }
            try { txtemail.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["a7"].Value.ToString(); }
            catch { }

            ///////////////////////////////nomrelerin oxunmasi ucun////////////////////////////////////////
            string commandText = "SELECT * FROM Telefon WHERE 1=1";
            
            try  { commandText += " and c1 like '%" + txtad.Text + "%'"; } catch { };
            MyData.selectCommand("baza.accdb",commandText);

            MyData.dtmaintelefon = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmaintelefon);

            try { tel1.Text = MyData.dtmaintelefon.Rows[0]["c2"].ToString(); }
            catch { };
            try { tel2.Text = MyData.dtmaintelefon.Rows[0]["c3"].ToString(); }
            catch { };

            tabControl1.SelectedTab = tabPage1;
        }

        private void dataGridView2_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            tabControl1.SelectedTab = tabPage3;

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (base.ClientSize.Height > 700) dataGridView2.Height = 2025/10; else dataGridView2.Height = 1925/10;

            button6.Text = "◄ ►"; 
            timer2.Enabled = false;
            if (dataGridView1.Height <= base.ClientSize.Height / 2.3) { button6.Text = "▼"; timer1.Enabled = false; return; }

            dataGridView1.Height -= 10;
            panel6.Top -= 10;
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            if (base.ClientSize.Height > 700) dataGridView2.Height = 2025/10; else dataGridView2.Height = 1925/10;

            button6.Text = "◄ ►";
            timer1.Enabled = false;
            if (dataGridView1.Height >= base.ClientSize.Height - 175) { button6.Text = "▲"; timer2.Enabled = false; return; }

            dataGridView1.Height += 10;
            panel6.Top += 10;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Height >= base.ClientSize.Height / 2) timer1.Enabled = true;
            if (dataGridView1.Height < base.ClientSize.Height / 2) timer2.Enabled = true; 
        }

        private void button7_Click(object sender, EventArgs e)
        {
            
            DialogResult result2 = MessageBox.Show("Davam etmək istəyirsiniz?", "?", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (result2 == DialogResult.No) { return; }
            Cursor.Current = Cursors.WaitCursor;

            try
            {
                File.Copy("Etibarname.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Etibarname - " + txtnomresi.Text.Substring(0,9).ToString() + ".xlsx", true);
            }
            catch { MessageBox.Show("Etibarname.xlsx tapılmadı."); }

            int azn = 0;
            string c = MyChange.ReqemToMetn(Convert.ToInt32(cbMuddet.Text));

            if (cbAyGun.Text == "ay")
            {
                azn = Convert.ToInt32(cbMuddet.Text) * 3;
            }

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Etibarname - " + txtnomresi.Text.Substring(0,9).ToString() + ".xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];

            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oSheet.Cells[1, 4] = txtseriyaAGL.Text;
            oSheet.Cells[3, 3] = txtTarix.Text;
            oSheet.Cells[5, 2] = txtBiz.Text;
            if (radioButton1.Checked == true) oSheet.Cells[3, 7] = "Şəhadətnaməsi əsasında " + radioButton1.Text + "-yə məxsus";
            if (radioButton2.Checked == true) oSheet.Cells[3, 7] = "Şəhadətnaməsi əsasında " + radioButton2.Text + "-yə məxsus";
            oSheet.Cells[7, 2] = txtUnvanimiz.Text;
            oSheet.Cells[9, 2] = txtEsas.Text;
            if (txtSexsiyyet.Text != "")
            {
                oSheet.Cells[12, 1] = "Şəxsiyyət vəsiqəsi:";
                oSheet.Cells[12, 4] = txtSexsiyyet.Text;
            }
            if (txtSuruculuk.Text != "")
            {
                oSheet.Cells[12, 1] = "Sürücülük vəsiqəsi:";
                oSheet.Cells[12, 4] = txtSuruculuk.Text;
            }
            oSheet.Cells[14, 1] = txtUnvan.Text;
            try
            {
                if (txtUnvan.Text.Length > 50)
                {
                    for (int y = 30; y < txtUnvan.Text.Length; y++)
                    {
                        if (txtUnvan.Text.Substring(y, 1) == ",")
                        {
                            oSheet.Cells[14, 1] = txtUnvan.Text.Substring(0, y + 1);
                            oSheet.Cells[16, 1] = txtUnvan.Text.Substring(y + 2, txtUnvan.Text.Length - y - 2);
                            y = txtUnvan.Text.Length;
                        }

                    }
                }
            }
            catch { }
            oSheet.Cells[18, 1] = txtSAA.Text;
            oSheet.Cells[23, 1] = txtTexpasstarix.Text;
            oSheet.Cells[3, 3] = txtTarix.Text;
            oSheet.Cells[3, 3] = txtTarix.Text;


            oSheet.Cells[1, 7] = txttexpassnomre.Text;
            oSheet.Cells[5, 7] = txtmarka.Text;
            oSheet.Cells[7, 10] = txtmuherrik.Text;
            if (txtBan.Text != "") { oSheet.Cells[8, 8] = txtBan.Text; oSheet.Cells[8, 7] = lbban.Text; }
            if (txtSassi.Text != "") { oSheet.Cells[8, 8] = txtSassi.Text; oSheet.Cells[8, 7] = lbsassi.Text; }

            oSheet.Cells[9, 8] = txtBuraxilis.Text;
            oSheet.Cells[11, 8] = txtnomresi.Text;
            oSheet.Cells[20, 11] = "'" + txtEtibarnameBitme.Text;
            if (c != "") oSheet.Cells[19, 9] = cbMuddet.Text + " (" + c + ") " + cbAyGun.Text;
            else oSheet.Cells[19, 9] = cbMuddet.Text + " " + c + " " + cbAyGun.Text;

            oSheet.Cells[20, 14] = "  - ci";
            try { if (txtEtibarnameBitme.Text.Substring(txtEtibarnameBitme.Text.Length - 1, 1) == "3" || txtEtibarnameBitme.Text.Substring(txtEtibarnameBitme.Text.Length - 1, 1) == "4") { oSheet.Cells[20, 14] = "  - cü"; } }
            catch { }
            try { if (txtEtibarnameBitme.Text.Substring(txtEtibarnameBitme.Text.Length - 1, 1) == "6") { oSheet.Cells[20, 14] = "  - cı"; } }
            catch { }
            try { if (txtEtibarnameBitme.Text.Substring(txtEtibarnameBitme.Text.Length - 1, 1) == "9") { oSheet.Cells[20, 14] = "  - cu"; } }
            catch { }

            //  oSheet.PrintOut();
            //   oWB.Close(SaveChanges: false);
            // oXL.Application.Quit();

            try
            {

                if (cbCOPY.Checked == true)
                {
                    MyData.selectCommand("baza.accdb", "Select * From etibarnamearxiv where a1='" + txtseriyaAGL.Text + "'");
                    MyData.dtmainArxiv = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmainArxiv);

                    if (MyData.dtmainArxiv.Rows.Count > 0)
                    {
                        MyData.updateCommand("baza.accdb", "UPDATE etibarnamearxiv SET "
                                                                                                               + "a1 = '" + txtseriyaAGL.Text + "',"
                                                                                                               + "a2 = '" + txtTarix.Text + "',"
                                                                                                               + "a3 = '" + txtBiz.Text.Substring(0, 14).ToString() + "',"
                                                                                                               + "a4 = '" + txtEsas.Text + "',"
                                                                                                               + "a5 = '" + txtSexsiyyet.Text + "',"
                                                                                                               + "a6 = '" + txtSuruculuk.Text + "',"
                                                                                                               + "a7 = '" + txtUnvan.Text + "',"
                                                                                                               + "a8 = '" + txtSAA.Text + "',"
                                                                                                               + "a9 = '" + txtTexpasstarix.Text + "',"
                                                                                                               + "a10 = '" + txttexpassnomre.Text + "',"
                                                                                                               + "a11 = '" + txtmarka.Text + "',"
                                                                                                               + "a12 = '" + txtmuherrik.Text + "',"
                                                                                                               + "a13 = '" + txtBan.Text + "',"
                                                                                                               + "a14 = '" + txtSassi.Text + "',"
                                                                                                               + "a15 = '" + txtBuraxilis.Text + "',"
                                                                                                               + "a16 = '" + txtnomresi.Text + "',"
                                                                                                               + "a17 = '" + txtEtibarnameBitme.Text + "',"
                                                                                                               + "a18 = '" + azn + "' where a1=" + "'" + txtseriyaAGL.Text + "'");

                        arxivrefresh();
                    }

                    else
                    {
                        MyData.insertCommand("baza.accdb", "INSERT INTO etibarnamearxiv (a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13,a14,a15,a16,a17,a18)values("


                                                                                                            + "'" + txtseriyaAGL.Text + "',"
                                                                                                            + "'" + txtTarix.Text + "',"
                                                                                                            + "'" + txtBiz.Text.Substring(0, 14).ToString() + "',"
                                                                                                            + "'" + txtEsas.Text + "',"
                                                                                                            + "'" + txtSexsiyyet.Text + "',"
                                                                                                            + "'" + txtSuruculuk.Text + "',"
                                                                                                            + "'" + txtUnvan.Text + "',"
                                                                                                            + "'" + txtSAA.Text + "',"
                                                                                                            + "'" + txtTexpasstarix.Text + "',"
                                                                                                            + "'" + txttexpassnomre.Text + "',"
                                                                                                            + "'" + txtmarka.Text + "',"
                                                                                                            + "'" + txtmuherrik.Text + "',"
                                                                                                            + "'" + txtBan.Text + "',"
                                                                                                            + "'" + txtSassi.Text + "',"
                                                                                                            + "'" + txtBuraxilis.Text + "',"
                                                                                                            + "'" + txtnomresi.Text + "',"
                                                                                                            + "'" + txtEtibarnameBitme.Text + "',"
                                                                                                            + "'" + azn + "')");
                        arxivrefresh();

                    }

                    MyData.updateCommand("baza.accdb", "UPDATE etibarnamenomre  SET a1 ='" + txtseriyaAGL.Text + "'");
                    etibarnamenomre();

                    lbetibarnamemeblegi.Visible = false;
                    lbetibarnamemeblegi.Text = "";
                }
            }
            catch { MessageBox.Show("Kopyalama alinmadı"); };


            oXL.Visible = false;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
            oSheet.PrintOut();
            oXL.Visible = false;
            oWB.Close(SaveChanges: true);
            oXL.Application.Quit();

            //emeliyyatlar ucun
            MyData.insertCommand("baza.accdb", "insert into Emeliyyatlar (a1, a2, a3) Values ('" + DateTime.Now + "', 'ETİBARNAMƏ - " + txtnomresi.Text + " - " + txtSAA.Text + " - " + cbMuddet.Text + " " + cbAyGun.Text + "','" + Environment.MachineName + "')");
            VerilmisSonEtibarname();
            Cursor.Current = Cursors.Default;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            progressBar1.Value = 0;
            progressBar1.Maximum = 100;
            progressBar1.Step = 25;
            progressBar2.Value = 0;
            progressBar2.Maximum = 100;
            progressBar2.Step = 25;

            textBox1.Text = "";
            myrefresh();        progressBar1.PerformStep(); progressBar2.PerformStep();
            suruculer();        progressBar1.PerformStep(); progressBar2.PerformStep();
            arxivrefresh();     progressBar1.PerformStep(); progressBar2.PerformStep();
            etibarnamenomre();  progressBar1.PerformStep(); progressBar2.PerformStep();

            //Xeberler yenilensin deye---------------------

            lbxeberler.Text = dataGridView1.Rows.Count.ToString() + " nəqliyyat vasitəsi və " + dataGridView2.Rows.Count.ToString() + " sürücü qeydiyyata alınıb.";
            txtarxivmelumat.Text = dataGridView1.Rows.Count.ToString() + " nəqliyyat vasitəsi və " + dataGridView2.Rows.Count.ToString() + " sürücü qeydiyyata alınmışdır.";
            lbxeberler.Left = base.Width;
            timer4.Enabled = true; 
            Cursor.Current = Cursors.Default;
        }

        private void label3_Click(object sender, EventArgs e)
        {
            if (txtseriyaAGL.Enabled == false) { txtseriyaAGL.Enabled = true; return; }
            if (txtseriyaAGL.Enabled == true) { txtseriyaAGL.Enabled = false; return; }
        }

        private void timer3_Tick(object sender, EventArgs e)
        {
            DateTime dt = DateTime.Now;
            label43.Text = "Tarix: " + dt.ToShortDateString() + "          Saat: " + dt.ToLongTimeString();
            label43.Left = base.Width/2 - label43.Width/2;

            if (dt.Second > 0 && dt.Second < 10) label43.ForeColor = Color.Tomato; 
            if (dt.Second > 10 && dt.Second < 20) label43.ForeColor = Color.Red; 
            if (dt.Second > 20 && dt.Second < 30) label43.ForeColor = Color.Green; 
            if (dt.Second > 30 && dt.Second < 40) label43.ForeColor = Color.Blue;
            if (dt.Second > 40 && dt.Second < 50) label43.ForeColor = Color.Maroon;
            if (dt.Second > 50 && dt.Second < 60) label43.ForeColor = Color.SteelBlue;


            if (dt.Second > 0 && dt.Second < 10)  button6.BackColor = Color.Tomato;
            if (dt.Second > 10 && dt.Second < 20) button6.BackColor = Color.Red;
            if (dt.Second > 20 && dt.Second < 30) button6.BackColor = Color.Green;
            if (dt.Second > 30 && dt.Second < 40) button6.BackColor = Color.Blue;
            if (dt.Second > 40 && dt.Second < 50) button6.BackColor = Color.Maroon;
            if (dt.Second > 50 && dt.Second < 60) button6.BackColor = Color.SteelBlue;
        }

        private void label10_Click(object sender, EventArgs e)
        {
            if (txtEtibarnameBitme.Enabled == false) { txtEtibarnameBitme.Enabled = true; return; }
            if (txtEtibarnameBitme.Enabled == true) { txtEtibarnameBitme.Enabled = false; return; }
        }

        private void label5_Click(object sender, EventArgs e)
        {
            if (txtBiz.Enabled == false) { txtBiz.Enabled = true; return; }
            if (txtBiz.Enabled == true) { txtBiz.Enabled = false; return; }
        }

        private void label6_Click(object sender, EventArgs e)
        {
            if (txtUnvanimiz.Enabled == false) { txtUnvanimiz.Enabled = true; return; }
            if (txtUnvanimiz.Enabled == true) { txtUnvanimiz.Enabled = false; return; }
        }

        private void label33_Click(object sender, EventArgs e)
        {
            if (txtTexpasstarix.Enabled == false) { txtTexpasstarix.Enabled = true; return; }
            if (txtTexpasstarix.Enabled == true) { txtTexpasstarix.Enabled = false; return; }
        }

        private void label8_Click(object sender, EventArgs e)
        {
            if (txttexpassnomre.Enabled == false) { txttexpassnomre.Enabled = true; return; }
            if (txttexpassnomre.Enabled == true) { txttexpassnomre.Enabled = false; return; }
        }

        private void label34_Click(object sender, EventArgs e)
        {
            if (txtmarka.Enabled == false) { txtmarka.Enabled = true; return; }
            if (txtmarka.Enabled == true) { txtmarka.Enabled = false; return; }
        }

        private void label9_Click(object sender, EventArgs e)
        {
            if (txtmuherrik.Enabled == false) { txtmuherrik.Enabled = true; return; }
            if (txtmuherrik.Enabled == true) { txtmuherrik.Enabled = false; return; }
        }

        private void lbban_Click(object sender, EventArgs e)
        {
            if (txtBan.Enabled == false) { txtBan.Enabled = true; return; }
            if (txtBan.Enabled == true) { txtBan.Enabled = false; return; }
        }

        private void lbsassi_Click(object sender, EventArgs e)
        {
            if (txtSassi.Enabled == false) { txtSassi.Enabled = true; return; }
            if (txtSassi.Enabled == true) { txtSassi.Enabled = false; return; }
        }

        private void label37_Click(object sender, EventArgs e)
        {
            if (txtBuraxilis.Enabled == false) { txtBuraxilis.Enabled = true; return; }
            if (txtBuraxilis.Enabled == true) { txtBuraxilis.Enabled = false; return; }
        }

        private void label7_Click(object sender, EventArgs e)
        {
            if (txtEsas.Enabled == false) { txtEsas.Enabled = true; return; }
            if (txtEsas.Enabled == true) { txtEsas.Enabled = false; return; }
        }

        private void label30_Click(object sender, EventArgs e)
        {
            if (txtSexsiyyet.Enabled == false) { txtSexsiyyet.Enabled = true; return; }
            if (txtSexsiyyet.Enabled == true) { txtSexsiyyet.Enabled = false; return; }
        }

        private void label31_Click(object sender, EventArgs e)
        {
            if (txtSuruculuk.Enabled == false) { txtSuruculuk.Enabled = true; return; }
            if (txtSuruculuk.Enabled == true) { txtSuruculuk.Enabled = false; return; }
        }

        private void label28_Click(object sender, EventArgs e)
        {
            if (txtUnvan.Enabled == false) { txtUnvan.Enabled = true; return; }
            if (txtUnvan.Enabled == true) { txtUnvan.Enabled = false; return; }
        }

        private void label11_Click(object sender, EventArgs e)
        {
            if (cbAyGun.Enabled == true) { cbAyGun.Enabled = false; return; }
                cbAyGun.Enabled = true;
        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    MyData.selectCommand("baza.accdb", "SELECT * FROM etibarnamearxiv WHERE a16 like '%" + textBox2.Text + "%' or a8 like '%" + textBox2.Text + "%' or a2 like '%" + textBox2.Text + "%'");
                    MyData.dtmainArxiv= new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmainArxiv);
                    dataGridView3.DataSource = MyData.dtmainArxiv;
                }
                catch { };

            }

            int i, k=0;
            try
            {
                for (i = 0; i < MyData.dtmainArxiv.Rows.Count; i++) k += Convert.ToInt32(MyData.dtmainArxiv.Rows[i][18]);
            }
            catch { };
        }

        private void printToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                File.Copy("Etibarname.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Etibarname - " + txtnomre.Text + ".xlsx", true);
            }
            catch { MessageBox.Show("Etibarname.xlsx tapılmadı."); }

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Etibarname - " + txtnomre.Text + ".xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];

            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oSheet.Cells[1, 4] = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa1"].Value.ToString();
            oSheet.Cells[3, 3] = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa2"].Value.ToString();
            oSheet.Cells[5, 2] = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa3"].Value.ToString();
            oSheet.Cells[3, 7] = "Şəhadətnaməsi əsasında " + dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa3"].Value.ToString() + "-yə məxsus";
            oSheet.Cells[7, 2] = txtUnvanimiz.Text;
            oSheet.Cells[9, 2] = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa4"].Value.ToString();
            if (dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa5"].Value.ToString() != "")
            {
                oSheet.Cells[12, 1] = "Şəxsiyyət vəsiqəsi:";
                oSheet.Cells[12, 4] = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa5"].Value.ToString();
            }
            if (dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa6"].Value.ToString() != "")
            {
                oSheet.Cells[12, 1] = "Sürücülük vəsiqəsi:";
                oSheet.Cells[12, 4] = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa6"].Value.ToString();
            }
            oSheet.Cells[14, 1] = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa7"].Value.ToString();
            oSheet.Cells[18, 1] = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa8"].Value.ToString();
            oSheet.Cells[23, 1] = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa9"].Value.ToString();

            oSheet.Cells[1, 7] = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa10"].Value.ToString();
            oSheet.Cells[5, 7] = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa11"].Value.ToString();
            oSheet.Cells[7, 10] = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa12"].Value.ToString();
            if (dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa13"].Value.ToString() != "") { oSheet.Cells[8, 7] = "BAN №-si"; oSheet.Cells[8, 8] = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa13"].Value.ToString(); }
            if (dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa14"].Value.ToString() != "") { oSheet.Cells[8, 7] = "Şassi №-si"; oSheet.Cells[8, 8] = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa14"].Value.ToString(); }

            oSheet.Cells[9, 8] = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa15"].Value.ToString();
            oSheet.Cells[11, 8] = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa16"].Value.ToString();
            oSheet.Cells[20, 11] = "'" + dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa17"].Value.ToString();
            oSheet.Cells[19, 9] = "";

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
            //oSheet.PrintOut();
            //oWB.Close(SaveChanges: false);
            //oXL.Application.Quit();

        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e) // delete etibarname arxiv
        {
            string ST = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa0"].Value.ToString();

                try
                {
                    MyData.deleteCommand("baza.accdb", "DELETE FROM etibarnamearxiv WHERE Kod=" + ST);
                    arxivrefresh();
                } catch { };
        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                MyData.deleteCommand("baza.accdb", "DELETE FROM etibarnameneqliyyat WHERE c14=" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c14"].Value.ToString());
                myrefresh();
            }
            catch { }
        }

        private void cbMuddet_TextChanged(object sender, EventArgs e)
        {
            try
            {
                DateTime bugun = Convert.ToDateTime(dtEtibarnameBugun.Value);

                switch (cbAyGun.Text)
                {
                    case "ay":
                        dtEtibarnameBitme.Text = bugun.AddMonths(Convert.ToInt32(cbMuddet.Text)).ToShortDateString();
                        lbetibarnamemeblegi.Visible = true;
                        lbetibarnamemeblegi.Text = "* Ödəniləcək məbləğ - " + (3 * Convert.ToInt32(cbMuddet.Text)).ToString() + " AZN";
                        break;
                    case "gün":
                        dtEtibarnameBitme.Text = bugun.AddDays(Convert.ToInt32(cbMuddet.Text)).ToShortDateString();
                        lbetibarnamemeblegi.Visible = true;
                        lbetibarnamemeblegi.Text = "* Ödəniləcək məbləğ - " + Convert.ToInt32(cbMuddet.Text) / 10 + " AZN";
                        break;
                }
            }
            catch { }
        }

        private void label4_DoubleClick(object sender, EventArgs e)
        {
            if (txtTarix.Enabled == false) { txtTarix.Enabled = true; dtEtibarnameBugun.Enabled = true; return; }
            if (txtTarix.Enabled == true) { txtTarix.Enabled = false; dtEtibarnameBugun.Enabled = false; return; }
        }

        private void txtnomre1_KeyDown(object sender, KeyEventArgs e)
        {
            
            if (e.KeyCode == Keys.Enter)
            {
                txtbannomre2.Text = "";
                txtmuherriknomre3.Text = "";
                txtsassinomre4.Text = "";
                txttexpass5.Text = "";
                txtburaxilisili6.Text = "";
                txtpassverilmetarix7.Text = "";
                txtmarkasi8.Text = "";
                txtlizinqalanadi9.Text = "";
                txtlayihensi10.Text = "";
                txtzavodnomresi11.Text = "";
                txtreng12.Text = "";
                txtQeyd.Text = "";

                string commandText = "SELECT * FROM etibarnameneqliyyat WHERE 1=1";
                try
                {
                   commandText += " and c1 like '%" + txtnomre1.Text + "%'";
                }
                catch { };

                MyData.selectCommand("baza.accdb", commandText);
                MyData.dtmain = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmain);

                try { txtnomre1.Text = MyData.dtmain.Rows[0]["c1"].ToString(); } catch { };
                try { txtbannomre2.Text = MyData.dtmain.Rows[0]["c8"].ToString(); } catch { };
                try { txtmuherriknomre3.Text = MyData.dtmain.Rows[0]["c9"].ToString(); } catch { };
                try { txtsassinomre4.Text = MyData.dtmain.Rows[0]["c10"].ToString(); } catch { };
                try { txttexpass5.Text = MyData.dtmain.Rows[0]["c12"].ToString(); } catch { };
                try { txtburaxilisili6.Text = MyData.dtmain.Rows[0]["c6"].ToString(); } catch { };
                try { txtpassverilmetarix7.Text = MyData.dtmain.Rows[0]["c7"].ToString(); } catch { };
                try { txtmarkasi8.Text = MyData.dtmain.Rows[0]["c2"].ToString(); } catch { };
                try { txtlizinqalanadi9.Text = MyData.dtmain.Rows[0]["c3"].ToString(); } catch { };
                try { txtlayihensi10.Text = MyData.dtmain.Rows[0]["c4"].ToString(); } catch { };
                try { txtzavodnomresi11.Text = MyData.dtmain.Rows[0]["c11"].ToString(); } catch { };
                try { txtreng12.Text = MyData.dtmain.Rows[0]["c5"].ToString(); } catch { };
                try { txtQeyd.Text = MyData.dtmain.Rows[0]["c13"].ToString(); } catch { };
                myrefresh();
            }
        }

        private void txtad_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                    txtsexsiyyetseriya.Text = "";
                    txtsexsiyyetnomre.Text = "";
                    txtsurucuseriya.Text = "";
                    txtsurucunomre.Text = "";
                    txtunvansurucu.Text = "";
                    dttarix.Text = "";
                    txtemail.Text = "";
                    tel1.Text = "";
                    tel2.Text = "";

               

                string commandText =  "SELECT * FROM etibarnamesurucu WHERE 1=1";
                try
                {
                    commandText += " and a1 like '%" + txtad.Text + "%'";
                }
                catch { };
                MyData.selectCommand("baza.accdb", commandText);
                MyData.dtmainSuruculer = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainSuruculer);

                try { txtad.Text = MyData.dtmainSuruculer.Rows[0]["a1"].ToString(); }
                catch { };
                try { txtsexsiyyetseriya.Text = MyData.dtmainSuruculer.Rows[0]["a2"].ToString().Substring(0, 3); }
                catch { };
                try { txtsexsiyyetnomre.Text = MyData.dtmainSuruculer.Rows[0]["a2"].ToString().Substring(6, MyData.dtmainSuruculer.Rows[0][2].ToString().Length - 6); }
                catch { };
                try { txtsurucuseriya.Text = MyData.dtmainSuruculer.Rows[0]["a3"].ToString().Substring(0, 2); }
                catch { };
                try { txtsurucunomre.Text = MyData.dtmainSuruculer.Rows[0]["a3"].ToString().Substring(5, MyData.dtmainSuruculer.Rows[0][3].ToString().Length - 5); }
                catch { };
                try { txtunvansurucu.Text = MyData.dtmainSuruculer.Rows[0]["a4"].ToString(); }
                catch { };
                try { txtlayihe.Text = MyData.dtmainSuruculer.Rows[0]["a5"].ToString(); }
                catch { };
                try { dttarix.Text = MyData.dtmainSuruculer.Rows[0]["a6"].ToString(); }
                catch { };
                try { txtemail.Text = MyData.dtmainSuruculer.Rows[0]["a7"].ToString(); }
                catch { };
                try { tel1.Text = MyData.dtmainSuruculer.Rows[0]["a8"].ToString(); }
                catch { };
                try { tel2.Text = MyData.dtmainSuruculer.Rows[0]["a9"].ToString(); }
                catch { };

                commandText = "SELECT * FROM Telefon WHERE 1=1";
                try
                {
                    commandText += " and c1 like " + "'%" + txtad.Text + "%'";
                }
                catch { };
                MyData.selectCommand("baza.accdb", commandText);
                MyData.dtmaintelefon = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmaintelefon);

                try { tel1.Text = MyData.dtmaintelefon.Rows[0][1].ToString(); }
                catch { };
                try { tel2.Text = MyData.dtmaintelefon.Rows[0][2].ToString(); }
                catch { };

                suruculer();
            }
        }

        private void txtSAA_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Cursor.Current = Cursors.WaitCursor;
                progressBar1.Value = 0;
                progressBar1.Maximum = 100;
                progressBar1.Step = 50;

                btvesiqe.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);

                string commandText = "SELECT * FROM etibarnamesurucu WHERE 1=1";
                try
                {
                   commandText += " and a1 like " + "'%" + txtSAA.Text + "%'";
                }
                catch { }; progressBar1.PerformStep();
                MyData.selectCommand("baza.accdb",commandText);
                MyData.dtmainSuruculer= new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainSuruculer);
                dataGridView2.DataSource = MyData.dtmainSuruculer;
                VerilmisSonEtibarname(); progressBar1.PerformStep();
                Cursor.Current = Cursors.Default;
            }
        }

        private void txtnomresi_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Cursor.Current = Cursors.WaitCursor;
                progressBar1.Value = 0;
                progressBar1.Maximum = 100;
                progressBar1.Step = 50;
                try
                {
                    MyData.selectCommand("baza.accdb", "SELECT * FROM etibarnameneqliyyat WHERE c1 like '%" + txtnomresi.Text + "%'");

                    MyData.dtmain= new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmain);
                    dataGridView1.DataSource = MyData.dtmain;
                }
                catch { }; progressBar1.PerformStep();

                VerilmisSonEtibarname(); progressBar1.PerformStep(); 
                Cursor.Current = Cursors.Default;
            }
        }

        private void əlaqəToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Elaqe elaqe = new Elaqe();
            elaqe.ShowDialog();
        }

        private void timer4_Tick(object sender, EventArgs e)
        {
            lbxeberler.Left -= 2;
            if (lbxeberler.Left <= btMMX2.Right + 20) { timer4.Enabled = false; }
        }

        private void txtnomresi_TextChanged(object sender, EventArgs e)
        {
            btmmx.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
            btgecikme.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);

            //timer6.Enabled = true;
            btBorcMelumat.Text = "";

            try
            {
                MyData.selectCommand("baza.accdb", "SELECT * FROM MMX WHERE a2 like '" + txtnomresi.Text + "' and a9 like 'Xeyr'");
                //oledbadapter1.SelectCommand.CommandText += "and a9 like Xeyr";
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
                btmmx.ForeColor = Color.Red;
                if (tt != 0) btmmx.Text = "* MMX (" + MyData.dtmainMMX.Rows.Count.ToString() + ") - " + tt.ToString() + " AZN";
                if (tt == 0) btmmx.Text = "* MMX (" + MyData.dtmainMMX.Rows.Count.ToString() + ")";
                //btMMX.Visible = true;
                btMMX2.Text = "MMX (" + MyData.dtmainMMX.Rows.Count.ToString() + ")";
                btMMX2.ForeColor = Color.White;
                btMMX2.BackColor = Color.Red;
                btMMX2.FlatAppearance.BorderColor = Color.Red;
            }
            else
            {
                btmmx.ForeColor = Color.Green;
                btmmx.Text = "* MMX yoxdur";
                //btMMX.Visible = false;
                btMMX2.Text = "MMX";
                btMMX2.ForeColor = Color.Gray;
                btMMX2.BackColor = Color.Silver;
                btMMX2.FlatAppearance.BorderColor = Color.Gray;
            } 
            //Agbank yaxud AGLizinq Layihesi oldugunu secmek ucun
            /*if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c13"].Value.ToString() == "AGL") { radioButton1.Checked = true; radioButton2.Checked = false; }
            else { radioButton2.Checked = true; radioButton1.Checked = false; }*/

            if (txtnomresi.Text.Length > 8)
            {
                MyData.selectCommand("baza.accdb", "SELECT * FROM Etibarnameneqliyyat WHERE c1 like '" + txtnomresi.Text + "'");
                MyData.dtmain = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmain);

                //////gecikmenin yoxlanmasi ucun-----------------------------------------------------------------
                try
                {
                    DateTime dt = DateTime.Now;
                    string gecikme = "";

                    gecikme = MyData.dtmain.Rows[0]["c4"].ToString();

                    String name = "licschkre";
                    String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    "2.xlsx" +
                                    ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

                    OleDbConnection con = new OleDbConnection(constr);
                    OleDbCommand oconn = new OleDbCommand("Select Adı, Layihe, [V#K#Qalıq], [V#K#% məbləği], [Dəbbə məbləği], [Cərimə % məbləği], [Son əməl#tarixi], [K#p#b#tarixi], [Qalıq], [% məbləği] From [" + name + "$] WHERE Layihe Like '%" + gecikme + "%'", con);
                    con.Open();
                    OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                    DataTable data = new DataTable();
                    data.Clear();
                    sda.Fill(data);
                    con.Close();

                    //dateTimePicker1.Text = data.Rows[0]["K#p#b#tarixi"].ToString().Substring(0, 2) + "-" + dt.AddMonths(1).ToString().Substring(3, 2) + "-" + dt.Year;
                    btBorcMelumat.Text = "";
                    btBorcMelumat.Text = "AGLizinq (" + data.Rows[0][0].ToString() + ")" + Environment.NewLine + "Ümumi Gecikmə - " + Math.Round(Convert.ToDouble(data.Rows[0]["V#K#Qalıq"]) + Convert.ToDouble(data.Rows[0]["V#K#% məbləği"]) + Convert.ToDouble(data.Rows[0]["Dəbbə məbləği"]) + Convert.ToDouble(data.Rows[0]["Cərimə % məbləği"]), 2).ToString() + " AZN (Cərimə+dəbbə - " + Math.Round(Convert.ToDouble(data.Rows[0]["Dəbbə məbləği"]) + Convert.ToDouble(data.Rows[0]["Cərimə % məbləği"]), 2).ToString() + " AZN)" + Environment.NewLine + "Ümumi borc - " + Math.Round(Convert.ToDouble(data.Rows[0]["Qalıq"]) + Convert.ToDouble(data.Rows[0]["V#K#Qalıq"]) + Convert.ToDouble(data.Rows[0]["V#K#% məbləği"]) + Convert.ToDouble(data.Rows[0]["Dəbbə məbləği"]) + Convert.ToDouble(data.Rows[0]["Cərimə % məbləği"]) + Convert.ToDouble(data.Rows[0]["% məbləği"]), 2).ToString() + " AZN" + Environment.NewLine + "Son əməliyyat tarixi - " + data.Rows[0]["Son əməl#tarixi"].ToString() + Environment.NewLine + "Bağlanma tarixi - " + data.Rows[0]["K#p#b#tarixi"].ToString();

                    if (Math.Round(Convert.ToDouble(data.Rows[0][2]) + Convert.ToDouble(data.Rows[0][3]) + Convert.ToDouble(data.Rows[0][4]) + Convert.ToDouble(data.Rows[0][5]), 2) > 0)
                    {
                        btgecikme.ForeColor = Color.Red; btgecikme.Text = "* Gecikmə (" + (Math.Round(Convert.ToDouble(data.Rows[0][2]) + Convert.ToDouble(data.Rows[0][3]), 2) + Math.Round(Convert.ToDouble(data.Rows[0][4]) + Convert.ToDouble(data.Rows[0][5]), 2)).ToString() + " Azn)";
                    }
                    else { btgecikme.ForeColor = Color.Green; btgecikme.Text = "* Gecikmə yoxdur"; }

                    //Qeydleri oxumaq ucun------------------------------------
                    try
                    {
                        button14.Text = "";
                        string txt = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c3"].Value.ToString(), txt2 = "";
                        for (int i = 0; i < txt.Length; i++) if (txt.Substring(i, 1) == " ") { txt2 = txt.Substring(i + 1, 4); i = txt.Length - 1; }

                        MyData.selectCommand("baza.accdb", "SELECT * FROM Qeydler WHERE c2 like '%" + txt2.ToUpper(MyChange.DilDeyisme) + "%'");
                        MyData.dtmainsozverenler = new DataTable();
                        MyData.oledbadapter1.Fill(MyData.dtmainsozverenler);
                        button14.Text = "- " + Convert.ToDateTime(MyData.dtmainsozverenler.Rows[0][1]).ToShortDateString().ToString() + " TARİXİNDƏ " + MyData.dtmainsozverenler.Rows[0][2].ToString();
                    }
                    catch { }

                }
                catch
                {

                    try
                    {
                        string gecikme = "";

                        gecikme = MyData.dtmain.Rows[0]["c4"].ToString();

                        String name = "Кредитный портфель";
                        String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                        "1.xlsx" +
                                        ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

                        OleDbConnection con = new OleDbConnection(constr);
                        OleDbCommand oconn = new OleDbCommand("Select [Наименование клиента], [Номер контракта], [Остаток просрочки на дату в манатном эквиваленте], [Просроченные проценты в манатном эквиваленте], [Штрафные проценты в манатном эквиваленте], [Последняя дата погашения процентов (или просроченных процентов)] From [" + name + "$] WHERE [Номер контракта] Like '%" + gecikme + "%' and [Способ выдачи кредита] Like 'LEASING'", con);
                        con.Open();
                        OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                        DataTable dataBank = new DataTable();
                        dataBank.Clear();
                        sda.Fill(dataBank);
                        con.Close();

                        string txt2 = dataBank.Rows[0][0].ToString();
                        btBorcMelumat.Text = "";
                        btBorcMelumat.Text = "AGBank (" + dataBank.Rows[0][0].ToString() + ")" + Environment.NewLine + "Ümumi Gecikmiş borc - " + Math.Round(Convert.ToDouble(dataBank.Rows[0][2]) + Convert.ToDouble(dataBank.Rows[0][3]) + Convert.ToDouble(dataBank.Rows[0][4]), 2).ToString() + " AZN" + Environment.NewLine + "O cümlədən Dəbbə borc - " + Math.Round(Convert.ToDouble(dataBank.Rows[0][4]), 2).ToString() + " AZN" + Environment.NewLine + "Son əməliyyat tarixi - " + Convert.ToDateTime(dataBank.Rows[0][5]).ToShortDateString();

                        //gecikmenin siyahida yasil olmasi ucun checkboxda
                        if (Math.Round(Convert.ToDouble(dataBank.Rows[0][2]) + Convert.ToDouble(dataBank.Rows[0][3]) + Convert.ToDouble(dataBank.Rows[0][4]), 2) > 0)
                        {
                            btgecikme.ForeColor = Color.Red; btgecikme.Text = "* Gecikmə (" + (Math.Round(Convert.ToDouble(dataBank.Rows[0][2]) + Convert.ToDouble(dataBank.Rows[0][3]), 2) + Math.Round(Convert.ToDouble(dataBank.Rows[0][4]), 2)).ToString() + " Azn)";
                        }
                        else { btgecikme.ForeColor = Color.Green; btgecikme.Text = "* Gecikmə yoxdur"; }

                        //Qeydleri oxumaq ucun------------------------------------
                        try
                        {
                            button14.Text = "";
                            txt2 = txt2.Substring(0, 4); ////bank musterilerinin adinin ilk dord herfi

                            MyData.selectCommand("baza.accdb", "SELECT * FROM Qeydler WHERE c2 like '%" + txt2.ToUpper(MyChange.DilDeyisme) + "%'");
                            MyData.dtmainsozverenler = new DataTable();
                            MyData.oledbadapter1.Fill(MyData.dtmainsozverenler);
                            button14.Text = "- " + Convert.ToDateTime(MyData.dtmainsozverenler.Rows[0][1]).ToShortDateString().ToString() + " TARİXİNDƏ " + MyData.dtmainsozverenler.Rows[0][2].ToString();
                        }
                        catch { }
                    }
                    catch { btgecikme.Text = "* Gecikmə yoxdur"; btgecikme.ForeColor = Color.Green; btBorcMelumat.Text = "Portfeldə bu Layihə üzrə məlumat yoxdur."; radioButton2.Checked = true; button14.Text = ""; }

                }
            }  
        }

        private void btMMX2_Click(object sender, EventArgs e)
        {
            MMX mmx = new MMX();
            mmx.Show();
            
            if (btMMX2.BackColor==Color.Red) mmx.textBox1.Text = txtnomresi.Text;
        }

        private void telefonKitabçasıToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Telefon telefon = new Telefon();
            telefon.Show();
        }
        
        private void dataGridView3_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView3.EditMode = DataGridViewEditMode.EditProgrammatically;

            try
            {
                MyData.updateCommand("baza.accdb", "UPDATE etibarnamearxiv SET "
                                                                                     + "a1 ='" + dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa1"].Value.ToString() + "',"
                                                                                     + "a2 ='" + dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa2"].Value.ToString() + "',"
                                                                                     + "a3 ='" + dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa3"].Value.ToString() + "',"
                                                                                     + "a4 ='" + dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa4"].Value.ToString() + "',"
                                                                                     + "a5 ='" + dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa5"].Value.ToString() + "',"
                                                                                     + "a6 ='" + dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa6"].Value.ToString() + "',"
                                                                                     + "a7 ='" + dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa7"].Value.ToString() + "',"
                                                                                     + "a8 ='" + dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa8"].Value.ToString() + "',"
                                                                                     + "a9 ='" + dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa9"].Value.ToString() + "',"
                                                                                     + "a10 ='" + dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa10"].Value.ToString() + "',"
                                                                                     + "a11 ='" + dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa11"].Value.ToString() + "',"
                                                                                     + "a12 ='" + dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa12"].Value.ToString() + "',"
                                                                                     + "a13 ='" + dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa13"].Value.ToString() + "',"
                                                                                     + "a14 ='" + dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa14"].Value.ToString() + "',"
                                                                                     + "a15 ='" + dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa15"].Value.ToString() + "',"
                                                                                     + "a16 ='" + dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa16"].Value.ToString() + "',"
                                                                                     + "a17 ='" + dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa17"].Value.ToString() + "',"
                                                                                     + "a18 ='" + dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa18"].Value.ToString() + "'"
                                                                                     + " WHERE Kod Like'" + dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa0"].Value.ToString() + "'");
            }
            catch { MessageBox.Show("Əməliyyat baş tutmadı."); }
        }

        private void editToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            dataGridView3.EditMode = DataGridViewEditMode.EditOnEnter;
        }

        private void txtlayihensi10_TextChanged(object sender, EventArgs e)
        {
            txtlayihe.Text = txtlayihensi10.Text;
            txtlayihe.BackColor = Color.Yellow;
            txtlayihensi10.BackColor = Color.Yellow;
        }

        private void txtnomre1_MouseClick(object sender, MouseEventArgs e)
        {
            if (txtnomre1.Text == "Nömrə yaz Enter ilə yoxla") { txtnomre1.Text = ""; txtnomre1.ForeColor = Color.Black; }
            txtnomre1.ForeColor = Color.Black;
        }

        private void txtnomre1_Leave(object sender, EventArgs e)
        {
            if (txtnomre1.Text == "") { txtnomre1.Text = "Nömrə yaz Enter ilə yoxla"; txtnomre1.ForeColor = Color.Gray; }
        }

        private void button11_Click_1(object sender, EventArgs e)
        {
            Odenisler odenisler = new Odenisler();
            odenisler.Show();
            odenisler.radioButton3.Checked = true;
            odenisler.comboBox4.Text = txtnomresi.Text;
        }


        private void timer5_Tick(object sender, EventArgs e)
        {
            txtarxivmelumat.Left -= 1;
            if (txtarxivmelumat.Right < 0) txtarxivmelumat.Left = base.Width;
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            MiaGovAz miagovaz = new MiaGovAz();
            miagovaz.Show();
            //System.Diagnostics.Process.Start("http://mia.gov.az/?/az/driverlicense/");
        }

        private void label2_Click(object sender, EventArgs e)
        {
            if (label2.ForeColor == Color.Black) { label2.ForeColor = Color.Red; return; }
            label2.ForeColor = Color.Black;
        }

        private void button13_Click(object sender, EventArgs e)
        {
            DialogResult result2 = MessageBox.Show("Davam etmək istəyirsiniz?", "?", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (result2 == DialogResult.No) { return; }
            Cursor.Current = Cursors.WaitCursor;

            WordDoc();

            if (cbCOPY.Checked == true)
            {
                try
                {
                    MyData.insertCommand("baza.accdb", "INSERT INTO etibarnamearxiv (a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13,a14,a15,a16,a17,a18)values("


                                                                                                        + "'" + "Xarici" + "',"
                                                                                                        + "'" + txtTarix.Text + "',"
                                                                                                        + "'" + txtBiz.Text.Substring(0, 14).ToString() + "',"
                                                                                                        + "'" + txtEsas.Text + "',"
                                                                                                        + "'" + txtSexsiyyet.Text + "',"
                                                                                                        + "'" + txtSuruculuk.Text + "',"
                                                                                                        + "'" + txtUnvan.Text + "',"
                                                                                                        + "'" + txtSAA.Text + "',"
                                                                                                        + "'" + txtTexpasstarix.Text + "',"
                                                                                                        + "'" + txttexpassnomre.Text + "',"
                                                                                                        + "'" + txtmarka.Text + "',"
                                                                                                        + "'" + txtmuherrik.Text + "',"
                                                                                                        + "'" + txtBan.Text + "',"
                                                                                                        + "'" + txtSassi.Text + "',"
                                                                                                        + "'" + txtBuraxilis.Text + "',"
                                                                                                        + "'" + txtnomresi.Text + "',"
                                                                                                        + "'" + txtEtibarnameBitme.Text + "',"
                                                                                                        + "'" + (Convert.ToInt32(cbMuddet.Text) * 3).ToString() + "')");
                    arxivrefresh();
                }
                catch { }
            }

            string commandText = "";
            if (radioButton1.Checked == true) commandText = "insert into Emeliyyatlar (a1, a2, a3) Values ('" + DateTime.Now + "', 'ETIBARNAMƏ (BEYNƏLXALQ) - " + txtnomresi.Text + " - " + txtSAA.Text + " - " + cbMuddet.Text + " " + cbAyGun.Text + " (" + radioButton1.Text + ")','" + Environment.MachineName +"')";
            if (radioButton2.Checked == true) commandText = "insert into Emeliyyatlar (a1, a2, a3) Values ('" + DateTime.Now + "', 'ETIBARNAMƏ (BEYNƏLXALQ) - " + txtnomresi.Text + " - " + txtSAA.Text + " - " + cbMuddet.Text + " " + cbAyGun.Text + " (" + radioButton2.Text + ")','" + Environment.MachineName +"')";
            MyData.insertCommand("baza.accdb", commandText);
            VerilmisSonEtibarname(); 
            Cursor.Current = Cursors.Default;
        }

        private void checkboxCOPY_CheckedChanged(object sender, EventArgs e)
        {
            if (cbCOPY.Checked == true)  cbCOPY.ForeColor = Color.Green;
            if (cbCOPY.Checked == false) cbCOPY.ForeColor = Color.Red;
        }

        private void vesiqe_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage1;

            try { txtad.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["columname1"].Value.ToString(); }
            catch { }
            try { txtsexsiyyetseriya.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["a2"].Value.ToString().Substring(0, 3); }
            catch { };
            try { txtsexsiyyetnomre.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["a2"].Value.ToString().Substring(6, MyData.dtmainSuruculer.Rows[dataGridView2.CurrentCell.RowIndex][2].ToString().Length - 6); }
            catch { };
            try { txtsurucuseriya.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["a3"].Value.ToString().Substring(0, 2); }
            catch { };
            try { txtsurucunomre.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["a3"].Value.ToString().Substring(5, MyData.dtmainSuruculer.Rows[dataGridView2.CurrentCell.RowIndex][3].ToString().Length - 5); }
            catch { };
            try { txtunvansurucu.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["a4"].Value.ToString(); }
            catch { }
            try { txtlayihe.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["a5"].Value.ToString(); }
            catch { }
            try { dttarix.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["a6"].Value.ToString(); }
            catch { } 
            try { dttarix.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["Email"].Value.ToString(); }
            catch { }
        }

        private void btmmx_Click(object sender, EventArgs e)
        {
            MMX mmx = new MMX();
            mmx.Show();
            mmx.textBox1.Text = txtnomresi.Text;
        }

        private void btgecikme_Click(object sender, EventArgs e)
        {
            Odenisler odenisler = new Odenisler();
            odenisler.Show();
            odenisler.radioButton3.Checked = true;
            odenisler.comboBox4.Text = txtnomresi.Text;
            System.Windows.Forms.SendKeys.Send("{ENTER}");
        }

        private void timer6_Tick(object sender, EventArgs e)
        {
            if (btmmx.ForeColor == Color.Red) { btmmx.Font = new Font("Times New Roman", 14, FontStyle.Bold); timer6.Enabled = false; timer7.Enabled = true; } else btmmx.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
            if (btgecikme.ForeColor == Color.Red) { btgecikme.Font = new Font("Times New Roman", 14, FontStyle.Bold); timer6.Enabled = false; timer7.Enabled = true; } else btgecikme.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
            if (btvesiqe.ForeColor == Color.Red) { btvesiqe.Font = new Font("Times New Roman", 14, FontStyle.Bold); timer6.Enabled = false; timer7.Enabled = true; } else btvesiqe.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
        }

        private void timer7_Tick(object sender, EventArgs e)
        {
            if (btmmx.ForeColor == Color.Red) { btmmx.Font = new Font("Times New Roman", 10, FontStyle.Bold); ; timer7.Enabled = false; timer6.Enabled = true; }
            if (btgecikme.ForeColor == Color.Red) { btgecikme.Font = new Font("Times New Roman", 10, FontStyle.Bold); ; timer7.Enabled = false; timer6.Enabled = true; }
            if (btvesiqe.ForeColor == Color.Red) { btvesiqe.Font = new Font("Times New Roman", 10, FontStyle.Bold); ; timer7.Enabled = false; timer6.Enabled = true; }
        }

        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {
            if (btmmx.Text == "* MMX" || btmmx.Text == "* MMX yoxdur") { MessageBox.Show("Protokollar təhvil verilib."); return; }
            DateTime dt = DateTime.Now;
            string k=txtnomresi.Text;
            try
            {
                MyData.updateCommand("baza.accdb", "UPDATE MMX SET "
                                                                                     + "a8 ='Bəli',"
                                                                                     + "a9 ='Bəli " + dt.ToShortDateString() + "'"
                                                                                     + " WHERE NOT a9 Like '%Bəli%' and a2 Like '%" + txtnomresi.Text + "%'");
                MessageBox.Show("Protokollar təhvil verildi.");
                txtnomresi.Text = "";
                txtnomresi.Text = k;
            }
            catch { MessageBox.Show("Əməliyyat baş tutmadı."); }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            DialogResult result2 = MessageBox.Show("Davam etmək istəyirsiniz?", "?", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (result2 == DialogResult.No) { return; }

            Cursor.Current = Cursors.WaitCursor;

            WordDocHuquqi();

            if (cbCOPY.Checked == true)
            {
                try
                {
                    MyData.insertCommand("baza.accdb", "INSERT INTO etibarnamearxiv (a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13,a14,a15,a16,a17,a18)values("


                                                                                                        + "'" + "Yol vərəqəsi" + "',"
                                                                                                        + "'" + txtTarix.Text + "',"
                                                                                                        + "'" + txtBiz.Text.Substring(0, 14).ToString() + "',"
                                                                                                        + "'" + txtEsas.Text + "',"
                                                                                                        + "'" + txtSexsiyyet.Text + "',"
                                                                                                        + "'" + txtSuruculuk.Text + "',"
                                                                                                        + "'" + txtUnvan.Text + "',"
                                                                                                        + "'" + txtSAA.Text + "',"
                                                                                                        + "'" + txtTexpasstarix.Text + "',"
                                                                                                        + "'" + txttexpassnomre.Text + "',"
                                                                                                        + "'" + txtmarka.Text + "',"
                                                                                                        + "'" + txtmuherrik.Text + "',"
                                                                                                        + "'" + txtBan.Text + "',"
                                                                                                        + "'" + txtSassi.Text + "',"
                                                                                                        + "'" + txtBuraxilis.Text + "',"
                                                                                                        + "'" + txtnomresi.Text + "',"
                                                                                                        + "'" + txtEtibarnameBitme.Text + "',"
                                                                                                        + "'" + (Convert.ToInt32(cbMuddet.Text) * 3).ToString() + "')");
                    arxivrefresh();
                }
                catch { }
            }

            string commandText = "";
            if (radioButton1.Checked == true) commandText = "insert into Emeliyyatlar (a1, a2, a3) Values ('" + DateTime.Now + "', 'ETIBARNAMƏ (Huquqi şəxs) - " + txtnomresi.Text + " - " + txtSAA.Text + " - " + cbMuddet.Text + " " + cbAyGun.Text + " (" + radioButton1.Text + ")','" + Environment.MachineName + "')";
            if (radioButton2.Checked == true) commandText = "insert into Emeliyyatlar (a1, a2, a3) Values ('" + DateTime.Now + "', 'ETIBARNAMƏ (Huquqi şəxs) - " + txtnomresi.Text + " - " + txtSAA.Text + " - " + cbMuddet.Text + " " + cbAyGun.Text + " (" + radioButton2.Text + ")','" + Environment.MachineName + "')";
            MyData.selectCommand("baza.accdb", commandText);

            VerilmisSonEtibarname();
            Cursor.Current = Cursors.Default;
        }

        private void button16_Click(object sender, EventArgs e)
        {
            txtSAA.Text = dataGridView3.Rows[dataGridView3.Rows.Count-1].Cells["aa8"].Value.ToString();

            btvesiqe.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
           string commandText = "SELECT * FROM etibarnamesurucu WHERE 1=1";
            try
            {
                commandText += " and a1 like '%" + txtSAA.Text + "%'";
            }
            catch { };
            MyData.selectCommand("baza.accdb", commandText);

            MyData.dtmainSuruculer= new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainSuruculer);
            dataGridView2.DataSource = MyData.dtmainSuruculer;
        }

        private void dYPÜçünMelumatlarToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                File.Copy("Bos.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Neqliyyat - " + ".xlsx", true);
            }
            catch { MessageBox.Show("Bos.xlsx tapılmadı."); }

            int a, b, c;
            string ss = "", kk = "", tt = "";
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Neqliyyat - " + ".xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;

            oSheet.Cells[1, 1] = "Dövlət Nömrə Nişanı";
            oSheet.Cells[1, 2] = "Marka";
            oSheet.Cells[1, 3] = "Lizinq alan";
            oSheet.Cells[1, 4] = "Sürücülər";
            oSheet.Cells[1, 5] = "Telefon";

            MyData.selectCommand("baza.accdb", "SELECT * FROM etibarnamearxiv");
            MyData.dtmainArxiv= new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainArxiv);
            b = MyData.dtmainArxiv.Rows.Count;

            for (a = 0; a < dataGridView1.Rows.Count; a++)
            {
                MyData.selectCommand("baza.accdb", "SELECT * FROM Telefon WHERE c1 Like '%" + dataGridView1.Rows[a].Cells["c3"].Value.ToString() + "%'");
                MyData.dtmaintelefon= new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmaintelefon);

                    tt = "";
                    try
                    {
                        if (MyData.dtmaintelefon.Rows[0]["c2"].ToString() != "       -  -" && MyData.dtmaintelefon.Rows[0]["c2"].ToString() != "") tt = MyData.dtmaintelefon.Rows[0]["c2"].ToString();
                        if (MyData.dtmaintelefon.Rows[0]["c3"].ToString() != "       -  -" && MyData.dtmaintelefon.Rows[0]["c3"].ToString() != "") tt += Environment.NewLine + MyData.dtmaintelefon.Rows[0]["c3"].ToString();
                    }
                    catch { }

                    oSheet.Cells[a + 2, 1] = dataGridView1.Rows[a].Cells["c1"].Value.ToString();
                    oSheet.Cells[a + 2, 2] = dataGridView1.Rows[a].Cells["c2"].Value.ToString();
                    oSheet.Cells[a + 2, 3] = dataGridView1.Rows[a].Cells["c3"].Value.ToString() + Environment.NewLine + tt;

                    MyData.selectCommand("baza.accdb", "SELECT * FROM etibarnamearxiv WHERE a16 Like '%" + dataGridView1.Rows[a].Cells["c1"].Value.ToString() + "%' and Kod > " + (b / 2).ToString());
                    MyData.dtmainArxiv= new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmainArxiv);

                    for (c = 0; c < MyData.dtmainArxiv.Rows.Count; c++)
                    {
                        try
                        {
                            ss += MyData.dtmainArxiv.Rows[c]["a6"].ToString() + " - " + MyData.dtmainArxiv.Rows[c]["a8"].ToString() + " - " + MyData.dtmainArxiv.Rows[c]["a2"].ToString() + " - " + MyData.dtmainArxiv.Rows[c]["a17"].ToString() + Environment.NewLine;
                        }
                        catch { }

                        try
                    {
                            MyData.selectCommand("baza.accdb", "SELECT * FROM Telefon WHERE c1 Like '%" + MyData.dtmainArxiv.Rows[c]["a8"].ToString() + "%'");
                            MyData.dtmaintelefon= new DataTable();
                            MyData.oledbadapter1.Fill(MyData.dtmaintelefon);

                            if (MyData.dtmaintelefon.Rows[0]["c2"].ToString() != "" && MyData.dtmaintelefon.Rows[0]["c3"].ToString() != "" && MyData.dtmaintelefon.Rows[0]["c2"].ToString() != "       -  -" && MyData.dtmaintelefon.Rows[0]["c3"].ToString() != "       -  -") kk += MyData.dtmaintelefon.Rows[0]["c2"].ToString() + " / " + MyData.dtmaintelefon.Rows[0]["c3"].ToString() + Environment.NewLine;
                            if (MyData.dtmaintelefon.Rows[0]["c2"].ToString() != "" && MyData.dtmaintelefon.Rows[0]["c2"].ToString() != "       -  -" && MyData.dtmaintelefon.Rows[0]["c3"].ToString() == "       -  -" || MyData.dtmaintelefon.Rows[0]["c3"].ToString() == "") kk += MyData.dtmaintelefon.Rows[0]["c2"].ToString() + Environment.NewLine;
                            if (MyData.dtmaintelefon.Rows[0]["c2"].ToString() == "" && MyData.dtmaintelefon.Rows[0]["c3"].ToString() == "" || MyData.dtmaintelefon.Rows[0]["c2"].ToString() == "       -  -" && MyData.dtmaintelefon.Rows[0]["c3"].ToString() == "       -  -") kk += "*" + Environment.NewLine;
                        }
                        catch { }
                    }

                    oSheet.Cells[a + 2, 4] = ss;
                    oSheet.Cells[a + 2, 5] = kk;
                    ss = ""; kk = "";

                oSheet.Range["A" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["B" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["C" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["D" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["E" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
            }

            oSheet.Range["A" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["B" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["C" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["D" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["E" + 1].Borders.LineStyle = Excel.Constants.xlSolid;

            //   oSheet.PrintOut();
            //  oWB.Close(SaveChanges: false);
            //  oXL.Workbooks.Close();
            // oXL.Application.Quit();
            //  oXL.Quit();
            oSheet.Columns.AutoFit();
            oSheet.Rows.AutoFit();
        }

        private void dataGridView3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                try
                {
                    string ST = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["aa0"].Value.ToString();

                    MyData.deleteCommand("baza.accdb", "DELETE FROM etibarnamearxiv WHERE Kod=" + ST);
                    arxivrefresh();
                }
                catch { }
            }
        }

        private void dateTimePicker2_ValueChanged_1(object sender, EventArgs e)
        {
            DateTime dt = Convert.ToDateTime(dtEtibarnameBugun.Value);
            txtTarix.Text = dt.Day + " " + MyChange.TarixSozle(dt) + " " + dt.Year;
        }

        private void aGLizinqToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBox1.Text = "AGL";
            myrefresh();

            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            try
            {
                File.Copy("Bos.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Neqliyyat - " + ".xlsx", true);
            }
            catch { MessageBox.Show("Bos.xlsx tapılmadı."); }

            int a, b, c;
            string ss = "", kk = "", tt = "";

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Neqliyyat - " + ".xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;

            oSheet.Cells[1, 1] = "Dövlət Nömrə Nişanı";
            oSheet.Cells[1, 2] = "Marka";
            oSheet.Cells[1, 3] = "Lizinq alan";
            oSheet.Cells[1, 4] = "Sürücülər";
            oSheet.Cells[1, 5] = "Telefon";

            MyData.selectCommand("baza.accdb", "SELECT * FROM etibarnamearxiv");
            MyData.dtmainArxiv= new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainArxiv);
            b = MyData.dtmainArxiv.Rows.Count;

            for (a = 0; a < dataGridView1.Rows.Count; a++)
            {
                MyData.selectCommand("baza.accdb", "SELECT * FROM Telefon WHERE c1 Like '%" + dataGridView1.Rows[a].Cells["c3"].Value.ToString() + "%'");
                MyData.dtmaintelefon= new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmaintelefon);

                tt = "";
                try
                {
                    if (MyData.dtmaintelefon.Rows[0]["c2"].ToString() != "       -  -" && MyData.dtmaintelefon.Rows[0]["c2"].ToString() != "") tt = MyData.dtmaintelefon.Rows[0]["c2"].ToString();
                    if (MyData.dtmaintelefon.Rows[0]["c3"].ToString() != "       -  -" && MyData.dtmaintelefon.Rows[0]["c3"].ToString() != "") tt += Environment.NewLine + MyData.dtmaintelefon.Rows[0]["c3"].ToString();
                }
                catch { }

                oSheet.Cells[a + 2, 1] = dataGridView1.Rows[a].Cells["c1"].Value.ToString();
                oSheet.Cells[a + 2, 2] = dataGridView1.Rows[a].Cells["c2"].Value.ToString();
                oSheet.Cells[a + 2, 3] = dataGridView1.Rows[a].Cells["c3"].Value.ToString() + Environment.NewLine + tt;

                MyData.selectCommand("baza.accdb", "SELECT * FROM etibarnamearxiv WHERE a16 Like '%" + dataGridView1.Rows[a].Cells["c1"].Value.ToString() + "%' and Kod > " + (b / 2).ToString());
                MyData.dtmainArxiv= new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainArxiv);

                for (c = 0; c < MyData.dtmainArxiv.Rows.Count; c++)
                {
                    try
                    {
                        ss = ss + MyData.dtmainArxiv.Rows[c]["a6"].ToString() + " - " + MyData.dtmainArxiv.Rows[c]["a8"].ToString() + " - " + MyData.dtmainArxiv.Rows[c]["a2"].ToString() + " - " + MyData.dtmainArxiv.Rows[c]["a17"].ToString() + Environment.NewLine;
                    }
                    catch { }

                    try
                    {
                        MyData.selectCommand("baza.accdb", "SELECT * FROM Telefon WHERE c1 Like '%" + MyData.dtmainArxiv.Rows[c]["a8"].ToString() + "%'");
                        MyData.dtmaintelefon= new DataTable();
                        MyData.oledbadapter1.Fill(MyData.dtmaintelefon);

                        if (MyData.dtmaintelefon.Rows[0]["c2"].ToString() != "" && MyData.dtmaintelefon.Rows[0]["c3"].ToString() != "" && MyData.dtmaintelefon.Rows[0]["c2"].ToString() != "       -  -" && MyData.dtmaintelefon.Rows[0]["c3"].ToString() != "       -  -") kk = kk + MyData.dtmaintelefon.Rows[0]["c2"].ToString() + " / " + MyData.dtmaintelefon.Rows[0]["c3"].ToString() + Environment.NewLine;
                        if (MyData.dtmaintelefon.Rows[0]["c2"].ToString() != "" && MyData.dtmaintelefon.Rows[0]["c2"].ToString() != "       -  -" && MyData.dtmaintelefon.Rows[0]["c3"].ToString() == "       -  -" || MyData.dtmaintelefon.Rows[0]["c3"].ToString() == "") kk = kk + MyData.dtmaintelefon.Rows[0]["c2"].ToString() + Environment.NewLine;
                        if (MyData.dtmaintelefon.Rows[0]["c2"].ToString() == "" && MyData.dtmaintelefon.Rows[0]["c3"].ToString() == "" || MyData.dtmaintelefon.Rows[0]["c2"].ToString() == "       -  -" && MyData.dtmaintelefon.Rows[0]["c3"].ToString() == "       -  -") kk = kk + "*" + Environment.NewLine;
                    }
                    catch { }
                }

                oSheet.Cells[a + 2, 4] = ss;
                oSheet.Cells[a + 2, 5] = kk;
                ss = ""; kk = "";

                oSheet.Range["A" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["B" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["C" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["D" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["E" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
            }

            oSheet.Range["A" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["B" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["C" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["D" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["E" + 1].Borders.LineStyle = Excel.Constants.xlSolid;

            //  oSheet.PrintOut();
            //  oWB.Close(SaveChanges: false);
            //  oXL.Workbooks.Close();
            //  oXL.Application.Quit();
            //  oXL.Quit();
            oSheet.Columns.AutoFit();
            oSheet.Rows.AutoFit();
        }

        private void aGBankToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBox1.Text = "AGB";
            myrefresh();

            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            try
            {
                File.Copy("Bos.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Neqliyyat - " + ".xlsx", true);
            }
            catch { MessageBox.Show("Bos.xlsx tapılmadı."); }

            int a, b, c;
            string ss = "", kk = "", tt = "";
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Neqliyyat - " + ".xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;


            oSheet.Cells[1, 1] = "Dövlət Nömrə Nişanı";
            oSheet.Cells[1, 2] = "Marka";
            oSheet.Cells[1, 3] = "Lizinq alan";
            oSheet.Cells[1, 4] = "Sürücülər";
            oSheet.Cells[1, 5] = "Telefon";

            MyData.selectCommand("baza.accdb", "SELECT * FROM etibarnamearxiv");
            MyData.dtmainArxiv = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainArxiv);

            b = MyData.dtmainArxiv.Rows.Count;

            for (a = 0; a < dataGridView1.Rows.Count; a++)
            {
                MyData.selectCommand("baza.accdb", "SELECT * FROM Telefon WHERE c1 Like '%" + dataGridView1.Rows[a].Cells["c3"].Value.ToString() + "%'");
                MyData.dtmaintelefon = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmaintelefon);

                tt = "";

                try
                {
                    if (MyData.dtmaintelefon.Rows[0]["c2"].ToString() != "       -  -" && MyData.dtmaintelefon.Rows[0]["c2"].ToString() != "") tt = MyData.dtmaintelefon.Rows[0]["c2"].ToString();
                    if (MyData.dtmaintelefon.Rows[0]["c3"].ToString() != "       -  -" && MyData.dtmaintelefon.Rows[0]["c3"].ToString() != "") tt += Environment.NewLine + MyData.dtmaintelefon.Rows[0]["c3"].ToString();
                }
                catch { }

                oSheet.Cells[a + 2, 1] = dataGridView1.Rows[a].Cells["c1"].Value.ToString();
                oSheet.Cells[a + 2, 2] = dataGridView1.Rows[a].Cells["c2"].Value.ToString();
                oSheet.Cells[a + 2, 3] = dataGridView1.Rows[a].Cells["c3"].Value.ToString() + Environment.NewLine + tt;

                MyData.selectCommand("baza.accdb", "SELECT * FROM etibarnamearxiv WHERE a16 Like '%" + dataGridView1.Rows[a].Cells["c1"].Value.ToString() + "%' and Kod > " + (b / 2).ToString());
                MyData.dtmainArxiv = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainArxiv);

                for (c = 0; c < MyData.dtmainArxiv.Rows.Count; c++)
                {
                    try
                    {
                        ss = ss + MyData.dtmainArxiv.Rows[c]["a6"].ToString() + " - " + MyData.dtmainArxiv.Rows[c]["a8"].ToString() + " - " + MyData.dtmainArxiv.Rows[c]["a2"].ToString() + " - " + MyData.dtmainArxiv.Rows[c]["a17"].ToString() + Environment.NewLine;
                    }
                    catch { }

                    try
                    {
                        MyData.selectCommand("baza.accdb", "SELECT * FROM Telefon WHERE c1 Like '%" + MyData.dtmainArxiv.Rows[c]["a8"].ToString() + "%'");
                        MyData.dtmaintelefon = new DataTable();
                        MyData.oledbadapter1.Fill(MyData.dtmaintelefon);

                        if (MyData.dtmaintelefon.Rows[0]["c2"].ToString() != "" && MyData.dtmaintelefon.Rows[0]["c3"].ToString() != "" && MyData.dtmaintelefon.Rows[0]["c2"].ToString() != "       -  -" && MyData.dtmaintelefon.Rows[0]["c3"].ToString() != "       -  -") kk = kk + MyData.dtmaintelefon.Rows[0]["c2"].ToString() + " / " + MyData.dtmaintelefon.Rows[0]["c3"].ToString() + Environment.NewLine;
                        if (MyData.dtmaintelefon.Rows[0]["c2"].ToString() != "" && MyData.dtmaintelefon.Rows[0]["c2"].ToString() != "       -  -" && MyData.dtmaintelefon.Rows[0]["c3"].ToString() == "       -  -" || MyData.dtmaintelefon.Rows[0]["c3"].ToString() == "") kk = kk + MyData.dtmaintelefon.Rows[0]["c2"].ToString() + Environment.NewLine;
                        if (MyData.dtmaintelefon.Rows[0]["c2"].ToString() == "" && MyData.dtmaintelefon.Rows[0]["c3"].ToString() == "" || MyData.dtmaintelefon.Rows[0]["c2"].ToString() == "       -  -" && MyData.dtmaintelefon.Rows[0]["c3"].ToString() == "       -  -") kk = kk + "*" + Environment.NewLine;

                    }
                    catch { }
                }

                oSheet.Cells[a + 2, 4] = ss;
                oSheet.Cells[a + 2, 5] = kk;
                ss = ""; kk = "";

                oSheet.Range["A" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["B" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["C" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["D" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["E" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
            }

            oSheet.Range["A" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["B" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["C" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["D" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["E" + 1].Borders.LineStyle = Excel.Constants.xlSolid;

            //   oSheet.PrintOut();
            //  oWB.Close(SaveChanges: false);
            //  oXL.Workbooks.Close();
            // oXL.Application.Quit();
            //  oXL.Quit();
            oSheet.Columns.AutoFit();
            oSheet.Rows.AutoFit();
        }

        private void deleteToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                MyData.deleteCommand("baza.accdb", "DELETE FROM etibarnamesurucu WHERE Kod=" + dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["Kod"].Value.ToString());
                suruculer();
            }
            catch { }
        }

        private void duzelişOlunmuşToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DYPUmumiMelumat();
        }

        private void btSon2_Click(object sender, EventArgs e)
        {
            try
            {
                txtSAA.Text = dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells["aa8"].Value.ToString();

                btvesiqe.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);

                string commandText = "SELECT * FROM etibarnamesurucu WHERE 1=1";
                try { commandText += " and a1 like " + "'%" + txtSAA.Text + "%'";} catch { };

                MyData.selectCommand("baza.accdb", commandText);
                MyData.dtmainSuruculer = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainSuruculer);
                dataGridView2.DataSource = MyData.dtmainSuruculer;
            }
            catch { }
        }

        private void btSon3_Click(object sender, EventArgs e)
        {
            try
            {
                txtSAA.Text = dataGridView3.Rows[dataGridView3.Rows.Count - 3].Cells["aa8"].Value.ToString();

                btvesiqe.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);

                string commandText = "SELECT * FROM etibarnamesurucu WHERE 1=1";
                try { commandText += " and a1 like " + "'%" + txtSAA.Text + "%'"; } catch { };

                MyData.selectCommand("baza.accdb", commandText);
                MyData.dtmainSuruculer = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainSuruculer);
                dataGridView2.DataSource = MyData.dtmainSuruculer;
            }
            catch { }
        
        }

        private void DtEtibarnameBitme_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    DateTime bugun = Convert.ToDateTime(dtEtibarnameBugun.Value).Date;
                    DateTime bitme = Convert.ToDateTime(dtEtibarnameBitme.Value).Date;

                    try { cbMuddet.Text = (Convert.ToInt32((bitme - bugun).Days)).ToString(); } catch { }

                    try { lbetibarnamemeblegi.Text = "* Ödəniləcək məbləğ - " + Convert.ToInt32(cbMuddet.Text) / 10 + " AZN"; } catch { }

                    cbAyGun.Text = "gün";
                }
            }
            catch { }
        }

        private void cbAyGun_TextChanged(object sender, EventArgs e)
        {
            try
            {
                DateTime bugun = Convert.ToDateTime(dtEtibarnameBugun.Value);

                switch (cbAyGun.Text)
                {
                    case "ay":
                        dtEtibarnameBitme.Text = bugun.AddMonths(Convert.ToInt32(cbMuddet.Text)).ToShortDateString();
                        lbetibarnamemeblegi.Visible = true;
                        lbetibarnamemeblegi.Text = "* Ödəniləcək məbləğ - " + (3 * Convert.ToInt32(cbMuddet.Text)).ToString() + " AZN";
                        break;
                    case "gün":
                        dtEtibarnameBitme.Text = bugun.AddDays(Convert.ToInt32(cbMuddet.Text)).ToShortDateString();
                        lbetibarnamemeblegi.Visible = true;
                        lbetibarnamemeblegi.Text = "* Ödəniləcək məbləğ - " + Convert.ToInt32(cbMuddet.Text) / 10 + " AZN";
                        break;
                }
            }
            catch { }
        }

        private void DtEtibarnameBitme_ValueChanged(object sender, EventArgs e)
        {
            try
            {
                DateTime bitme = Convert.ToDateTime(dtEtibarnameBitme.Value).Date;
                txtEtibarnameBitme.Text = bitme.Day + " " + MyChange.TarixSozle(bitme) + " " + bitme.Year;
                try { lbetibarnamemeblegi.Text = "* Ödəniləcək məbləğ - " + Convert.ToInt32(cbMuddet.Text) / 10 + " AZN"; } catch { }

                DateTime VesiqeBitme = Convert.ToDateTime(dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["a6"].Value).Date;
                if (bitme <= VesiqeBitme) { btvesiqe.ForeColor = Color.Green; return; }
                else { btvesiqe.ForeColor = Color.Red; return; }
            }
            catch { btvesiqe.ForeColor = Color.Red; }
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }
    }
}