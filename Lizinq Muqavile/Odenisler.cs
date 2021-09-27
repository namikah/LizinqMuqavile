using Nsoft;
using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Lizinq_Muqavile
{
    public partial class Odenisler : Form
    {

        Excel.Application oXL;
        Excel._Workbook oWB;
        Excel._Worksheet oSheet;

        public Odenisler()
        {
            InitializeComponent();
        }
        private void myrefresh()           //------------------------------------------------------------------------------
        {
            try
            {
                if (radioButton1.Checked)
                {
                    progressBar1.Value = 25;

                    String name = "licschkre";
                    String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    "2.xlsx" +
                                    ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

                    OleDbConnection con = new OleDbConnection(constr);
                    OleDbCommand oconn = new OleDbCommand("Select [Ö#G#], Adı, Layihe, [Lizinqin məbləği], [Qalıq], [V#K#Qalıq], [V#K#% məbləği], [% məbləği], [Dəbbə məbləği], [Cərimə % məbləği], [Qrafikda olan məbləğ], [Verilmə tarixi], [K#p#b#tarixi], [Son əməl#tarixi], [Lizinq hesabı] From [" + name + "$] WHERE [Ö#G#] Like '%" + comboBox4.Text + "%'", con);
                    con.Open();
                    OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                    DataTable data = new DataTable();
                    data.Clear();
                    sda.Fill(data);///////////////////datatableni sortlamaq ucun
                    progressBar1.Value = 60;
                    data.DefaultView.Sort = "Layihe desc";
                    data = data.DefaultView.ToTable(true);
                    con.Close();

                    dataGridView1.DataSource = data;
                    progressBar1.Value = 100;
                }
                else if (radioButton2.Checked)
                {
                    progressBar1.Value = 25;

                    String name = "licschkre";
                    String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    "2.xlsx" +
                                    ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

                    OleDbConnection con = new OleDbConnection(constr);
                    OleDbCommand oconn = new OleDbCommand("Select [Ö#G#], Adı, Layihe, [Lizinqin məbləği], [Qalıq], [V#K#Qalıq], [V#K#% məbləği], [% məbləği], [Dəbbə məbləği], [Cərimə % məbləği], [Qrafikda olan məbləğ], [Verilmə tarixi], [K#p#b#tarixi], [Son əməl#tarixi], [Lizinq hesabı]  From [" + name + "$] WHERE Adı Like '%" + comboBox4.Text + "%' or Layihe Like '%" + comboBox4.Text + "%'", con);
                    con.Open();
                    OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                    DataTable data = new DataTable();
                    data.Clear();
                    sda.Fill(data);///////////////////datatableni sortlamaq ucun
                    progressBar1.Value = 60;
                    data.DefaultView.Sort = "Layihe desc";
                    data = data.DefaultView.ToTable(true);
                    con.Close();

                    dataGridView1.DataSource = data;
                    progressBar1.Value = 100;
                }
                else if (radioButton3.Checked)
                {
                    progressBar1.Value = 25;
                    MyData.selectCommand("baza.accdb", "SELECT * FROM etibarnameneqliyyat WHERE c1 like '%" + comboBox4.Text + "%'");
                    MyData.dtmain = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmain);
                    comboBox4.Text = MyData.dtmain.Rows[0]["c1"].ToString();
                    comboBox4.Items.Clear();

                    for (int say = 0; say < MyData.dtmain.Rows.Count; say++) { comboBox4.Items.Add(MyData.dtmain.Rows[say]["c1"].ToString()); }
                    
                    String name = "licschkre";
                    String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    "2.xlsx" +
                                    ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

                        OleDbConnection con = new OleDbConnection(constr);
                        OleDbCommand oconn = new OleDbCommand("Select [Ö#G#], Adı, Layihe, [Lizinqin məbləği], [Qalıq], [V#K#Qalıq], [V#K#% məbləği], [% məbləği], [Dəbbə məbləği], [Cərimə % məbləği], [Qrafikda olan məbləğ], [Verilmə tarixi], [K#p#b#tarixi], [Son əməl#tarixi], [Lizinq hesabı]  From [" + name + "$] WHERE Layihe Like '%" + MyData.dtmain.Rows[0]["c4"].ToString() + "%'", con);
                       
                        con.Open();
                        OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                        DataTable data = new DataTable();
                        data.Clear();
                        sda.Fill(data);///////////////////datatableni sortlamaq ucun
                        progressBar1.Value = 60;
                        data.DefaultView.Sort = "Layihe desc";
                        data = data.DefaultView.ToTable(true);
                        con.Close();
                        dataGridView1.DataSource = data;
                        progressBar1.Value = 100;
                }
            }
            catch { }
        }

        private void odenisler_Load(object sender, EventArgs e)
        {
            String name = "licschkre";
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            "2.xlsx" +
                            ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(constr);
            OleDbCommand oconn = new OleDbCommand("Select Overdueday  From [" + name + "$]", con);
            con.Open();
            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
            DataTable data = new DataTable();
            data.Clear();
            sda.Fill(data);
            con.Close();
            base.Text += " - " + data.Rows[0][0].ToString();
            button7.Text = "Yenilənmə tarixi - " + data.Rows[0][0].ToString();

            MyChange.SetKeyboardLayout(MyChange.GetInputLanguageByName("AZ"));
            myrefresh();
        }

        private void çıxışToolStripMenuItem_Click(object sender, EventArgs e)
        {
            base.Close();
        }

        private void infoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Help help = new Help();
            help.ShowDialog();
        }

        private void əlaqəToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Elaqe elaqe = new Elaqe();
            elaqe.ShowDialog();
        }

        private void telefonKitabçasıToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Telefon telefon = new Telefon();
            telefon.Show();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            comboBox4.Text = DateTime.Now.ToShortDateString().Substring(0, 2).ToString();
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            comboBox4.Text = "";
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            
            //dataGridView1.DefaultCellStyle.ForeColor = Color.Red;
            //dataGridView1.Rows[0].DefaultCellStyle.BackColor = Color.Beige;

            richTextBox1.Text = "";
            listBox1.Items.Clear();
            listBox2.Items.Clear();
            button11.Text = "";
            button1.Text = "";
            button2.Text = "";
            button3.Text = "";
            button4.Text = "";
            button5.Text = "";
            button6.Text = "";

            try
            {
                string commandText = "SELECT * FROM etibarnamesurucu WHERE a5 like " + "'%" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Layihe"].Value.ToString().Substring(1, dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Layihe"].Value.ToString().Length - 1) + "%'";
                MyData.selectCommand("baza.accdb", commandText);
                MyData.dtmain = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmain);
            }
            catch { }

            try
            {
                for (int k = 0; k < MyData.dtmain.Rows.Count; k++)
                {
                  //button2.Text += MyData.dtmain.Rows[k][1].ToString()+Environment.NewLine;
                    listBox1.Items.Add(MyData.dtmain.Rows[k][1].ToString());
                }
            }
            catch { }

            if (listBox1.Items.Count > 0) listBox1.SetSelected(0, true); 



            ////////borclar ucun/////////////gecikmesi odenisi ve s.
            try
            {
                button11.Text = "Aylıq ödəniş -  " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Qrafikda olan məbləğ"].Value.ToString();
                button2.Text = "Gecikmə -  " + (Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["V#K#Qalıq"].Value) + Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["V#K#% məbləği"].Value)).ToString();
                if (Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["V#K#Qalıq"].Value) + Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["V#K#% məbləği"].Value) > 0) { button2.ForeColor = Color.DarkRed; button6.ForeColor = Color.DarkRed; } else { button2.ForeColor = Color.DarkGreen; button6.ForeColor = Color.DarkGreen; }
                button3.Text = "Cərimə+Dəbbə - " + (Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Dəbbə məbləği"].Value) + Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Cərimə % məbləği"].Value)).ToString();
                if (Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Dəbbə məbləği"].Value) + Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Cərimə % məbləği"].Value) > 0) { button3.ForeColor = Color.DarkRed; button6.ForeColor = Color.DarkRed; } else { button3.ForeColor = Color.DarkGreen;}
                button1.Text = "Son əməliyyat - " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Son əməl#tarixi"].Value.ToString();
                button4.Text = "Ümumi Borc -  " + (Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["V#K#Qalıq"].Value) + Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["V#K#% məbləği"].Value) + Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Dəbbə məbləği"].Value) + Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Cərimə % məbləği"].Value) + Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Qalıq"].Value) + Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["% məbləği"].Value)).ToString();
                button6.Text = "Ümumi Gecikmə -  " + (Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["V#K#Qalıq"].Value) + Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["V#K#% məbləği"].Value) + Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Dəbbə məbləği"].Value) + Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Cərimə % məbləği"].Value)).ToString();

            }
            catch { }


            /////////// qeydlerin load olunmasi ucun
            try
            {
                string txt = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Adı"].Value.ToString(), txt2 = "";
                for (int i = 0; i < txt.Length; i++) if (txt.Substring(i, 1) == " ") { txt2 = txt.Substring(i + 1, 4); i = txt.Length - 1; }
                
                
                
                MyData.selectCommand("baza.accdb", "SELECT * FROM Qeydler WHERE c2 like '%" + txt2 + "%'");
                MyData.dtmain= new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmain);
                richTextBox1.Text = Convert.ToDateTime(MyData.dtmain.Rows[0][1].ToString()).ToShortDateString() + " TARİXİNDƏ " + MyData.dtmain.Rows[0][2].ToString();
                
            }
            catch { }

            try //masinin markasini dartib getirmek ucun
            {
                
                
                
                MyData.selectCommand("baza.accdb", "SELECT * FROM Etibarnameneqliyyat WHERE c4 like " + "'%" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Layihe"].Value.ToString().Substring(1, dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Layihe"].Value.ToString().Length-1) + "%'");
                MyData.dtmain= new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmain);
                button5.Text = MyData.dtmain.Rows[0]["c1"].ToString() + "  /  " + MyData.dtmain.Rows[0]["c2"].ToString();
            }
            catch { }

            try //masinin markasini dartib getirmek ucun
            {
                MyData.selectCommand("baza.accdb", "SELECT * FROM PortfelStatus WHERE Layihe like '%" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Layihe"].Value.ToString().Substring(1, dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Layihe"].Value.ToString().Length - 1) + "%'");
                MyData.dtmain= new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmain);
                button8.Text = MyData.dtmain.Rows[0][3].ToString();
                if (button8.Text == "M") button8.Text = "Mehkeme";
                else if (button8.Text == "L") button8.Text = "Lizinq";
                else if (button8.Text == "P") button8.Text = "Problemli";
            }
            catch { }

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                MyData.selectCommand("baza.accdb", "SELECT * FROM Telefon WHERE c1 like '%" + listBox1.Items[listBox1.SelectedIndex].ToString() + "%'");
                MyData.dtmain= new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmain);
                listBox2.Items.Clear();
                listBox2.Items.Add(MyData.dtmain.Rows[0][1].ToString());
                listBox2.Items.Add(MyData.dtmain.Rows[0][2].ToString());
            }
            catch { }
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            comboBox4.Text = "";
        }

        private void etibarnaməYazToolStripMenuItem_Click(object sender, EventArgs e)
        {
            EtibarnameEsas etibarname = new EtibarnameEsas();
            etibarname.Show();


            etibarname.textBox1.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Layihe"].Value.ToString().Substring(1, dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Layihe"].Value.ToString().Length-1);
            etibarname.label2.ForeColor = Color.Red;
            System.Windows.Forms.SendKeys.Send("{Enter}");
        }

        private void ödənişQəbziToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int s = 0;
            string gecikme = (Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["V#K#Qalıq"].Value) + Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["V#K#% məbləği"].Value) + Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Dəbbə məbləği"].Value) + Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Cərimə % məbləği"].Value)).ToString();

            Qəbz qebz = new Qəbz();
            qebz.Show();

            qebz.txtteyinat.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Layihe"].Value.ToString() + " lizinq ödənişi";
            qebz.txtlizinqalan.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Adı"].Value.ToString();
            qebz.txtodeyen.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Adı"].Value.ToString();
            try { qebz.txthesabnomresi.Text = "5430" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Lizinq hesabı"].Value.ToString().Substring(5, dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Lizinq hesabı"].Value.ToString().Length - 8); }
            catch { }

            qebz.txt1.Text = gecikme;
            if (gecikme == "0") qebz.txt1.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Qrafikda olan məbləğ"].Value.ToString();
            {
                s = 0;
                for (int i = 0; i < dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Qrafikda olan məbləğ"].Value.ToString().Length; i++)
                {
                    if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Qrafikda olan məbləğ"].Value.ToString().Substring(i, 1) == ".") s = i;
                }

                try { if (s == 0) qebz.txt1.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Qrafikda olan məbləğ"].Value.ToString(); } catch { }
            }
        }

        private void ödənişQəbziToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Qəbz qebz = new Qəbz();
            qebz.Show();
        }

        public void saglam_olanlar()
        {
            string gecikme = "";
            DateTime dt = DateTime.Now;

            try { File.Copy("New Emphty.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Gecikme - " + dt.ToShortDateString() + ".doc", true); }
            catch { MessageBox.Show("'\\192.168.10.5\\Common\\AGLizinq\\New Emphty.doc' tapılmadı."); }

            object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\Gecikme - " + dt.ToShortDateString() + ".doc";

            string tarix = "", OdenisGunu = "";

            String name = "licschkre";
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            "2.xlsx" +
                            ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(constr);
            OleDbCommand oconn = new OleDbCommand("Select Adı, [V#K#Qalıq], [V#K#% məbləği], [Dəbbə məbləği], [Cərimə % məbləği], [Qrafikda olan məbləğ], [Verilmə tarixi], [overdueday], [Qalıq], [% məbləği], Layihe, [K#p#b#tarixi], [Son əməl#tarixi], [Uzadılma tarixi] From [" + name + "$] WHERE Not Qalıq=0 and [V#K#Qalıq]=0 and [V#K#% məbləği]=0 and [Cərimə % məbləği]=0 and [Dəbbə məbləği]=0", con);
            con.Open();
            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
            DataTable data1 = new DataTable();
            data1.Clear();
            sda.Fill(data1);
            con.Close();

            int k = 0;
            double CemiQaliq = 0;


            for (int i = 0; i < data1.Rows.Count; i++)
            {
                if (data1.Rows[i]["Uzadılma tarixi"].ToString() != "") OdenisGunu = data1.Rows[i]["Uzadılma tarixi"].ToString().Substring(0, 2);
                else OdenisGunu = data1.Rows[i]["K#p#b#tarixi"].ToString().Substring(0, 2);

                CemiQaliq += Math.Round(Convert.ToDouble(data1.Rows[i]["Qalıq"]) + Convert.ToDouble(data1.Rows[i]["% məbləği"]), 2);

                if ((Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"]) == 0))
                {
                    gecikme += Environment.NewLine + OdenisGunu + " ♦ " + data1.Rows[i]["Adı"].ToString() + " - " + Math.Round((Convert.ToDouble(data1.Rows[i]["% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"])), 2).ToString() + " ♦ aylıq - " + data1.Rows[i]["Qrafikda olan məbləğ"].ToString() + " ♦ " + data1.Rows[i]["Layihe"].ToString().Substring(1, data1.Rows[i]["Layihe"].ToString().Length - 1) + " ♦ son ödəmə - " + data1.Rows[i]["Son əməl#tarixi"].ToString();
                    k += 1;
                }


            }

            try
            {
                String name2 = "licschkre";
                String constr2 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                "2.xlsx" +
                                ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

                OleDbConnection con2 = new OleDbConnection(constr2);
                OleDbCommand oconn2 = new OleDbCommand("Select overdueday From [" + name2 + "$]", con2);
                con2.Open();
                OleDbDataAdapter sda2 = new OleDbDataAdapter(oconn2);
                DataTable data2 = new DataTable();
                data2.Clear();
                sda2.Fill(data2);
                con2.Close();
                tarix = "PORTFEL TARİXİ ♦ " + data2.Rows[0]["overdueday"].ToString();
            }
            catch { }

            Microsoft.Office.Interop.Word._Application oWord;
            object oMissing = Type.Missing;
            oWord = new Microsoft.Office.Interop.Word.Application();
            oWord.Visible = true;
            oWord.Documents.Open(FileName);
            oWord.Selection.TypeText(tarix + Environment.NewLine + gecikme + Environment.NewLine + Environment.NewLine + "İNFO:" + Environment.NewLine + "Sağlam müştərilər - " + data1.Rows.Count.ToString() + " ədəd, məbləğ - " + CemiQaliq + " azn");
            oWord.ActiveDocument.Save();
            //oWord.Quit();
            //MessageBox.Show("The text is inserted.");

            //}
            //catch { MessageBox.Show("Səhv var."); }

        }

        public void gecikmede_olanlar()
        {
            string gecikme = "";
            DateTime dt = DateTime.Now;

            try { File.Copy("New Emphty.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Gecikme - " + dt.ToShortDateString() + ".doc", true); }
            catch { MessageBox.Show("'\\192.168.10.5\\Common\\AGLizinq\\New Emphty.doc' tapılmadı."); }

            object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\Gecikme - " + dt.ToShortDateString() + ".doc";

            string tarix = "", OdenisGunu = "";
            //try
            //{

            String name = "licschkre";
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            "2.xlsx" +
                            ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(constr);
            OleDbCommand oconn = new OleDbCommand("Select Adı, [V#K#Qalıq], [V#K#% məbləği], [Dəbbə məbləği], [Cərimə % məbləği], [Qrafikda olan məbləğ], [Verilmə tarixi], [overdueday], [Qalıq], [% məbləği], Layihe, [K#p#b#tarixi], [Son əməl#tarixi], [Uzadılma tarixi] From [" + name + "$] WHERE NOT Layihe Like '%S-01/15-048/15%' and NOT Layihe Like '%S-032/14%' and Layihe Like '%S-%' or Layihe Like '%A-%'  or Layihe like '%126/08%' or Layihe like '%106/08%' or Layihe like '%158%' or Layihe like '%079%' or Layihe like '%174%' or Layihe like '%105/08%' or Layihe like '%105/08%' or Layihe like '%198%' or Layihe like '%193%' or Layihe like '%161%' or Layihe like '%196%' or Layihe like '%114%' or Layihe like '%L-01/16%' or Layihe like '%055%' or Layihe like '%01/13-167/10%'or Layihe like '%02/14-167/10%'or Layihe like '%03/14-167/10%'or Layihe like '%04/16-167/10%'", con);
            con.Open();
            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
            DataTable data1 = new DataTable();
            data1.Clear();
            sda.Fill(data1);
            con.Close();

            int k = 0, k2 = 0;
            double CemiGecikme = 0, Gecikme100 = 0, CemiQaliq = 0, NisbetFaiz = 0;


            for (int i = 0; i < data1.Rows.Count; i++)
            {
                if (data1.Rows[i]["Uzadılma tarixi"].ToString() != "") OdenisGunu = data1.Rows[i]["Uzadılma tarixi"].ToString().Substring(0, 2);
                else OdenisGunu = data1.Rows[i]["K#p#b#tarixi"].ToString().Substring(0, 2);

                CemiQaliq += Math.Round(Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"]) + Convert.ToDouble(data1.Rows[i]["Qalıq"]) + Convert.ToDouble(data1.Rows[i]["% məbləği"]), 2);

                if ((Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"]) > 0))
                {
                    ////ABDULLAYEV TOFIQ UCUN XUSUSI YAZILIM GECIKMEDE FAIZ BORCUN GORUNMEMESI UCUN.........................
                    if (data1.Rows[i]["Layihe"].ToString() == "'S-008/14") { gecikme += Environment.NewLine + OdenisGunu + " ♦ " + data1.Rows[i]["Adı"].ToString() + " - " + Math.Round((Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"])), 2).ToString() + "/" + Math.Round((Convert.ToDouble(data1.Rows[i]["% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"])), 2).ToString() + " ♦ aylıq - " + data1.Rows[i]["Qrafikda olan məbləğ"].ToString() + " ♦ " + data1.Rows[i]["Layihe"].ToString().Substring(1, data1.Rows[i]["Layihe"].ToString().Length - 1) + " ♦ son ödəmə - " + data1.Rows[i]["Son əməl#tarixi"].ToString(); }
                    else { gecikme += Environment.NewLine + OdenisGunu + " ♦ " + data1.Rows[i]["Adı"].ToString() + " - " + Math.Round((Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"])), 2).ToString() + "/" + Math.Round((Convert.ToDouble(data1.Rows[i]["% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"])), 2).ToString() + " ♦ aylıq - " + data1.Rows[i]["Qrafikda olan məbləğ"].ToString() + " ♦ " + data1.Rows[i]["Layihe"].ToString().Substring(1, data1.Rows[i]["Layihe"].ToString().Length - 1) + " ♦ son ödəmə - " + data1.Rows[i]["Son əməl#tarixi"].ToString(); }
                    
                    k += 1;
                    CemiGecikme += Math.Round(Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"]), 2);
                }

                if (Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"]) > 0 && Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"]) < 100)
                {
                    k2 += 1;
                    Gecikme100 += Math.Round(Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"]), 2);
                }
            }
            //notifyIcon1.Icon = SystemIcons.;
            NisbetFaiz = Math.Round(CemiGecikme / CemiQaliq * 100, 2);   //gecikmenin faizinin tapilmasi
            //pertfel vaxti ucun
            try
            {
                String name2 = "licschkre";
                String constr2 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                "2.xlsx" +
                                ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

                OleDbConnection con2 = new OleDbConnection(constr2);
                OleDbCommand oconn2 = new OleDbCommand("Select overdueday From [" + name2 + "$]", con2);
                con2.Open();
                OleDbDataAdapter sda2 = new OleDbDataAdapter(oconn2);
                DataTable data2 = new DataTable();
                data2.Clear();
                sda2.Fill(data2);
                con2.Close();
                tarix = "PORTFEL TARİXİ ♦ " + data2.Rows[0]["overdueday"].ToString();
            }
            catch { }

            string qeydler = "";
            
            
            
            MyData.selectCommand("baza.accdb", "Select * from Qeydler");
            MyData.dtmain= new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmain);

            if (MyData.dtmain.Rows.Count != 0)
            {
                for (int i = 0; i < MyData.dtmain.Rows.Count; i++)
                {
                    qeydler += Environment.NewLine + MyData.dtmain.Rows[i][1].ToString().Substring(0, 10) + " - " + MyData.dtmain.Rows[i][2].ToString();
                }

                qeydler = "QEYDLƏR:" + qeydler;
            }

            Microsoft.Office.Interop.Word._Application oWord;
            object oMissing = Type.Missing;
            oWord = new Microsoft.Office.Interop.Word.Application();
            oWord.Visible = true;
            oWord.Documents.Open(FileName);
            oWord.Selection.TypeText(tarix + Environment.NewLine + gecikme + Environment.NewLine + Environment.NewLine + "İNFO:" + Environment.NewLine + "Real müştərilər - " + data1.Rows.Count.ToString() + " ədəd, məbləğ - " + CemiQaliq + " azn" + Environment.NewLine + "Gecikmədə olanlar - " + k.ToString() + " ədəd, məbləğ - " + CemiGecikme + " azn  (Gecikmə faizlə - " + NisbetFaiz.ToString() + "%)" + Environment.NewLine + "O cümlədən, 100 AZN-ə kimi gecikənlər - " + k2.ToString() + " ədəd, məbləğ - " + Gecikme100 + " azn" + Environment.NewLine + Environment.NewLine + qeydler);
            oWord.ActiveDocument.Save();
            //oWord.Quit();
            //MessageBox.Show("The text is inserted.");

            //}
            //catch { MessageBox.Show("Səhv var."); }

        }

        private void aktivLizinqLayihələrToolStripMenuItem_Click(object sender, EventArgs e)
        {
            String name = "licschkre";
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            "2.xlsx" +
                            ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(constr);
            OleDbCommand oconn = new OleDbCommand("Select [Ö#G#], [Layihe], Adı, [Lizinqin məbləği], [Qalıq], [% məbləği], [V#K#Qalıq], [V#K#% məbləği], [Cərimə % məbləği], [Dəbbə məbləği], [Son əməl#tarixi], [Qrafikda olan məbləğ], [overdueday] From [" + name + "$] where NOT Qalıq + [V#K#Qalıq] + [% məbləği] + [V#K#% məbləği] + [Cərimə % məbləği] + [Dəbbə məbləği] = '0' and NOT Layihe Like '%S-009/14%' and NOT Layihe Like '%S-01/15-039/14%' and NOT Layihe Like '%S-01/15-048/15%' and NOT Layihe Like '%S-027/14%' and NOT Layihe Like '%S-050/15%' and NOT Layihe Like '%S-032/14%' and Layihe like '%S-%' or Layihe like '%L-01/14-126/08%' or Layihe like '%A-%' or Layihe like '%01/13-106/08%' or Layihe like '%L-158%' or Layihe like '%L-079%' or Layihe like '%L-174%' or Layihe like '%01/08-105/08%' or Layihe like '%02/08-105/08%' or Layihe like '%L-193%' or Layihe like '%L-196%' or Layihe like '%L-114%' or Layihe like '%L-01/13-167/10%' or Layihe like '%L-01/16%' or Layihe like '%L-195/11%'", con);
            con.Open();

            OleDbDataAdapter abc = new OleDbDataAdapter(oconn);
            DataTable dataABC = new DataTable();
            dataABC.Clear();
            abc.Fill(dataABC);
            dataABC.DefaultView.Sort = "Layihe desc";
            dataABC = dataABC.DefaultView.ToTable(true);
            //dataGridView1.DataSource = data;
            con.Close();

            try{
            int s, k, a = 0;

            try { File.Copy("Portfel.xlsm", "C:\\Users\\" + Environment.UserName + "\\Desktop\\PORTFEL Aktiv.xlsm", true); }
            catch { MessageBox.Show("'Portfel.xlsm' tapılmadı."); }

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\PORTFEL Aktiv.xlsm"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];

            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
            oSheet.PageSetup.Zoom = 80;
            //oSheet.Cells.EntireColumn.NumberFormat = "@";

                for (s = 0; s < dataABC.Columns.Count; s++)
                {
                    a = s + 2;

                    oSheet.Cells[1, a] = dataABC.Columns[s].ColumnName;
                    //oSheet.Cells[1, 4] = dataABC.Columns[s].ColumnName;
                    oSheet.Cells[1, 1] = "Nö";

                    for (k = 0; k < dataABC.Rows.Count; k++)
                    {
                        oSheet.Cells[k + 2, a] = dataABC.Rows[k][s].ToString();
                        if (dataABC.Columns[s].ColumnName == "Son əməl#tarixi") oSheet.Cells[k + 2, a] = "'" + dataABC.Rows[k][s].ToString();
                        //oSheet.Cells[k + 1, s + 1].Borders.LineStyle = Excel.Constants.xlSolid;

                        if (k != dataABC.Rows.Count) oSheet.Cells[k + 2, 1] = k + 1;
                    }
                }

                DateTime today = DateTime.Now;
                oSheet.Cells[1, 4] = "Adı                                                                                       " + today.ToShortDateString();

                double cem1 = 0, cem2 = 0, cem3 = 0, cem4 = 0, cem5 = 0, cem6 = 0;
                for (k = 0; k < dataABC.Rows.Count; k++)
                {
                    cem1 += Convert.ToDouble(dataABC.Rows[k]["Qalıq"]);
                    cem2 += Convert.ToDouble(dataABC.Rows[k]["% məbləği"]);
                    cem3 += Convert.ToDouble(dataABC.Rows[k]["V#K#Qalıq"]);
                    cem4 += Convert.ToDouble(dataABC.Rows[k]["V#K#% məbləği"]);
                    cem5 += Convert.ToDouble(dataABC.Rows[k]["Cərimə % məbləği"]);
                    cem6 += Convert.ToDouble(dataABC.Rows[k]["Dəbbə məbləği"]);
                }

                oSheet.Cells[dataABC.Rows.Count + 2, 6] = cem1.ToString();
                oSheet.Cells[dataABC.Rows.Count + 2, 7] = cem2.ToString();
                oSheet.Cells[dataABC.Rows.Count + 2, 8] = cem3.ToString();
                oSheet.Cells[dataABC.Rows.Count + 2, 9] = cem4.ToString();
                oSheet.Cells[dataABC.Rows.Count + 2, 10] = cem5.ToString();
                oSheet.Cells[dataABC.Rows.Count + 2, 11] = cem6.ToString();
                oSheet.Cells[dataABC.Rows.Count + 2, 12] = (cem1 + cem2 + cem3 + cem4 + cem5 + cem6).ToString();
                oSheet.Range[oSheet.Cells[dataABC.Rows.Count + 2, 12], oSheet.Cells[dataABC.Rows.Count + 2, 14]].Merge();

                oSheet.Cells[dataABC.Rows.Count + 2, 6].Font.Bold = true;
                oSheet.Cells[dataABC.Rows.Count + 2, 7].Font.Bold = true;
                oSheet.Cells[dataABC.Rows.Count + 2, 8].Font.Bold = true;
                oSheet.Cells[dataABC.Rows.Count + 2, 9].Font.Bold = true;
                oSheet.Cells[dataABC.Rows.Count + 2, 10].Font.Bold = true;
                oSheet.Cells[dataABC.Rows.Count + 2, 11].Font.Bold = true;
                oSheet.Cells[dataABC.Rows.Count + 2, 12].Font.Bold = true;
                oSheet.Columns.AutoFit();
                oSheet.Rows.AutoFit();

            oXL.DisplayAlerts = false;
            oWB.Save();
            //  oSheet.PrintOut();
            //  oWB.Close(SaveChanges: false);
            //  oXL.Application.Quit();
            }
            catch (System.Exception excep)
            {
                MessageBox.Show(excep.Message + " Məlumatlar Excel-ə doldurularkən, Excel-ə toxunmaq olmaz !!! ");
            }
        }

        private void sLayihələrToolStripMenuItem_Click(object sender, EventArgs e)
        {
            String name = "licschkre";
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            "2.xlsx" +
                            ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(constr);
            OleDbCommand oconn = new OleDbCommand("Select [Ö#G#], [Layihe], Adı, [Lizinqin məbləği], [Qalıq], [% məbləği], [V#K#Qalıq], [V#K#% məbləği], [Cərimə % məbləği], [Dəbbə məbləği], [Son əməl#tarixi], [Qrafikda olan məbləğ], [overdueday] From [" + name + "$] where Layihe like '%S%'", con);
            con.Open();

            OleDbDataAdapter abc = new OleDbDataAdapter(oconn);
            DataTable dataABC = new DataTable();
            dataABC.Clear();
            abc.Fill(dataABC);
            ///////////////////datatableni sortlamaq ucun
            dataABC.DefaultView.Sort = "Layihe desc";
            dataABC = dataABC.DefaultView.ToTable(true);
            //dataGridView1.DataSource = data;
            con.Close();

            try{
            int s, k, a = 0;

            try { File.Copy("Portfel.xlsm", "C:\\Users\\" + Environment.UserName + "\\Desktop\\PORTFEL (S).xlsm", true); }
            catch { MessageBox.Show("'Portfel.xlsm' tapılmadı."); }

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\PORTFEL (S).xlsm"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];

            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
            oSheet.PageSetup.Zoom = 80;
            //oSheet.Cells.EntireColumn.NumberFormat = "@";

            for (s = 0; s < dataABC.Columns.Count; s++)
            {
                a = s + 2;

                oSheet.Cells[1, a] = dataABC.Columns[s].ColumnName;
                oSheet.Cells[1, 1] = "Nö";

                for (k = 0; k < dataABC.Rows.Count; k++)
                {
                    oSheet.Cells[k + 2, a] = dataABC.Rows[k][s];
                    if (dataABC.Columns[s].ColumnName == "Son əməl#tarixi") oSheet.Cells[k + 2, a] = "'" + dataABC.Rows[k][s].ToString();
                    //oSheet.Cells[k + 1, s + 1].Borders.LineStyle = Excel.Constants.xlSolid;

                    if (k != dataABC.Rows.Count) oSheet.Cells[k + 2, 1] = k + 1;
                }
            }

                DateTime today = DateTime.Now;
                oSheet.Cells[1, 4] = "Adı                                                                                       " + today.ToShortDateString();

                double cem1 = 0, cem2 = 0, cem3 = 0, cem4 = 0, cem5 = 0, cem6 = 0;
            for (k = 0; k < dataABC.Rows.Count; k++)
            {
                cem1 += Convert.ToDouble(dataABC.Rows[k]["Qalıq"]);
                cem2 += Convert.ToDouble(dataABC.Rows[k]["% məbləği"]);
                cem3 += Convert.ToDouble(dataABC.Rows[k]["V#K#Qalıq"]);
                cem4 += Convert.ToDouble(dataABC.Rows[k]["V#K#% məbləği"]);
                cem5 += Convert.ToDouble(dataABC.Rows[k]["Cərimə % məbləği"]);
                cem6 += Convert.ToDouble(dataABC.Rows[k]["Dəbbə məbləği"]);
            }

            oSheet.Cells[dataABC.Rows.Count + 2, 6] = cem1.ToString();
            oSheet.Cells[dataABC.Rows.Count + 2, 7] = cem2.ToString();
            oSheet.Cells[dataABC.Rows.Count + 2, 8] = cem3.ToString();
            oSheet.Cells[dataABC.Rows.Count + 2, 9] = cem4.ToString();
            oSheet.Cells[dataABC.Rows.Count + 2, 10] = cem5.ToString();
            oSheet.Cells[dataABC.Rows.Count + 2, 11] = cem6.ToString();
            oSheet.Cells[dataABC.Rows.Count + 2, 12] = (cem1 + cem2 + cem3 + cem4 + cem5 + cem6).ToString();
            oSheet.Range[oSheet.Cells[dataABC.Rows.Count + 2, 12], oSheet.Cells[dataABC.Rows.Count + 2, 14]].Merge();

            oSheet.Cells[dataABC.Rows.Count + 2, 6].Font.Bold = true;
            oSheet.Cells[dataABC.Rows.Count + 2, 7].Font.Bold = true;
            oSheet.Cells[dataABC.Rows.Count + 2, 8].Font.Bold = true;
            oSheet.Cells[dataABC.Rows.Count + 2, 9].Font.Bold = true;
            oSheet.Cells[dataABC.Rows.Count + 2, 10].Font.Bold = true;
            oSheet.Cells[dataABC.Rows.Count + 2, 11].Font.Bold = true;
            oSheet.Cells[dataABC.Rows.Count + 2, 12].Font.Bold = true;
            oSheet.Columns.AutoFit();
            oSheet.Rows.AutoFit();

            oXL.DisplayAlerts = false;
            oWB.Save();
            //  oSheet.PrintOut();
            //  oWB.Close(SaveChanges: false);
            //  oXL.Application.Quit();
            }
            catch (System.Exception excep)
            {
                MessageBox.Show(excep.Message + " Məlumatlar Excel-ə doldurularkən, Excel-ə toxunmaq olmaz !!! ");
            }
        }

        private void lLayihələrToolStripMenuItem_Click(object sender, EventArgs e)
        {
            String name = "licschkre";
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            "2.xlsx" +
                            ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(constr);
            OleDbCommand oconn = new OleDbCommand("Select [Ö#G#], [Layihe], Adı, [Lizinqin məbləği], [Qalıq], [% məbləği], [V#K#Qalıq], [V#K#% məbləği], [Cərimə % məbləği], [Dəbbə məbləği], [Son əməl#tarixi], [Qrafikda olan məbləğ], [overdueday] From [" + name + "$] where Layihe like '%L%'", con);
            con.Open();

            OleDbDataAdapter abc = new OleDbDataAdapter(oconn);
            DataTable dataABC = new DataTable();
            dataABC.Clear();
            abc.Fill(dataABC);
            ///////////////////datatableni sortlamaq ucun
            dataABC.DefaultView.Sort = "Layihe desc";
            dataABC = dataABC.DefaultView.ToTable(true);
            //dataGridView1.DataSource = data;
            con.Close();

            try{
            int s, k, a = 0;

            try { File.Copy("Portfel.xlsm", "C:\\Users\\" + Environment.UserName + "\\Desktop\\PORTFEL (L).xlsm", true); }
            catch { MessageBox.Show("'Portfel.xlsm' tapılmadı."); }

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\PORTFEL (L).xlsm"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];

            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;

            //oSheet.Cells.EntireColumn.NumberFormat = "@";

            for (s = 0; s < dataABC.Columns.Count; s++)
            {
                a = s + 2;

                oSheet.Cells[1, a] = dataABC.Columns[s].ColumnName;
                oSheet.Cells[1, 1] = "Nö";

                for (k = 0; k < dataABC.Rows.Count; k++)
                {
                    oSheet.Cells[k + 2, a] = dataABC.Rows[k][s].ToString();
                    if (dataABC.Columns[s].ColumnName == "Son əməl#tarixi") oSheet.Cells[k + 2, a] = "'" + dataABC.Rows[k][s].ToString();
                    //oSheet.Cells[k + 1, s + 1].Borders.LineStyle = Excel.Constants.xlSolid;

                    if (k != dataABC.Rows.Count) oSheet.Cells[k + 2, 1] = k + 1;
                }
            }

                DateTime today = DateTime.Now;
                oSheet.Cells[1, 4] = "Adı                                                                                       " + today.ToShortDateString();

                double cem1 = 0, cem2 = 0, cem3 = 0, cem4 = 0, cem5 = 0, cem6 = 0;
            for (k = 0; k < dataABC.Rows.Count; k++)
            {
                cem1 += Convert.ToDouble(dataABC.Rows[k]["Qalıq"]);
                cem2 += Convert.ToDouble(dataABC.Rows[k]["% məbləği"]);
                cem3 += Convert.ToDouble(dataABC.Rows[k]["V#K#Qalıq"]);
                cem4 += Convert.ToDouble(dataABC.Rows[k]["V#K#% məbləği"]);
                cem5 += Convert.ToDouble(dataABC.Rows[k]["Cərimə % məbləği"]);
                cem6 += Convert.ToDouble(dataABC.Rows[k]["Dəbbə məbləği"]);
            }

            oSheet.Cells[dataABC.Rows.Count + 2, 6] = cem1.ToString();
            oSheet.Cells[dataABC.Rows.Count + 2, 7] = cem2.ToString();
            oSheet.Cells[dataABC.Rows.Count + 2, 8] = cem3.ToString();
            oSheet.Cells[dataABC.Rows.Count + 2, 9] = cem4.ToString();
            oSheet.Cells[dataABC.Rows.Count + 2, 10] = cem5.ToString();
            oSheet.Cells[dataABC.Rows.Count + 2, 11] = cem6.ToString();
            oSheet.Cells[dataABC.Rows.Count + 2, 12] = (cem1 + cem2 + cem3 + cem4 + cem5 + cem6).ToString();
            oSheet.Range[oSheet.Cells[dataABC.Rows.Count + 2, 12], oSheet.Cells[dataABC.Rows.Count + 2, 14]].Merge();

            oSheet.Cells[dataABC.Rows.Count + 2, 6].Font.Bold = true;
            oSheet.Cells[dataABC.Rows.Count + 2, 7].Font.Bold = true;
            oSheet.Cells[dataABC.Rows.Count + 2, 8].Font.Bold = true;
            oSheet.Cells[dataABC.Rows.Count + 2, 9].Font.Bold = true;
            oSheet.Cells[dataABC.Rows.Count + 2, 10].Font.Bold = true;
            oSheet.Cells[dataABC.Rows.Count + 2, 11].Font.Bold = true;
            oSheet.Cells[dataABC.Rows.Count + 2, 12].Font.Bold = true;
            oSheet.Columns.AutoFit();
            oSheet.Rows.AutoFit();

            oXL.DisplayAlerts = false;
            oWB.Save();
            //  oSheet.PrintOut();
            //  oWB.Close(SaveChanges: false);
            //  oXL.Application.Quit();
            }
            catch (System.Exception excep)
            {
                MessageBox.Show(excep.Message + " Məlumatlar Excel-ə doldurularkən, Excel-ə toxunmaq olmaz !!! ");
            }
        }

        private void lSLayihələrToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
        }

        private void bankLayihələrToolStripMenuItem_Click(object sender, EventArgs e)
        {
            String name = "Кредитный портфель";
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            "1.xlsx" +
                            ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(constr);
            OleDbCommand oconn = new OleDbCommand("Select [ID клиента], [Наименование клиента], [Дата заключения контракта], [Дата окончания контракта], [Валюта контракта], [Сумма контракта], [Сумма контракта в манатном эквиваленте], [Остаток основного долга на дату], [Остаток основного долга на дату в манатном эквиваленте], [Остаток просрочки на дату], [Остаток просрочки на дату в манатном эквиваленте], [Начисленные проценты], [Начисленные проценты в манатном эквиваленте], [Просроченные проценты], [Просроченные проценты в манатном эквиваленте], [Штрафные проценты], [Штрафные проценты в манатном эквиваленте], [Срок просрочки процентов], [Последняя дата погашения процентов (или просроченных процентов)], [Куратор кредита], [Способ выдачи кредита] From [" + name + "$]", con);
            con.Open();

            OleDbDataAdapter abc = new OleDbDataAdapter(oconn);
            DataTable dataABC = new DataTable();
            dataABC.Clear();
            abc.Fill(dataABC);
            ///////////////////datatableni sortlamaq ucun
            dataABC.DefaultView.Sort = "[Наименование клиента] ASC";
            dataABC = dataABC.DefaultView.ToTable(true);
            //dataGridView1.DataSource = data;
            con.Close();

            try
            {
                int s, k, a = 0;

                try { File.Copy("Portfel.xlsm", "C:\\Users\\" + Environment.UserName + "\\Desktop\\PORTFEL (AGBANK).xlsm", true); }
                catch { MessageBox.Show("'Portfel.xlsm' tapılmadı."); }

                //Get a new workbook.
                oXL = new Excel.Application();
                oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\PORTFEL (AGBANK).xlsm"));
                oSheet = (Excel._Worksheet)oWB.Sheets[1];

                oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                oXL.Visible = true;
                oSheet.Activate();
                oSheet.Range["A1"].Select();
                oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

                //oSheet.Cells.EntireColumn.NumberFormat = "@";

                for (s = 0; s < dataABC.Columns.Count; s++)
                {
                    a = s + 2;

                    oSheet.Cells[1, a] = dataABC.Columns[s].ColumnName;
                    oSheet.Cells[1, 1] = "Nö";

                    for (k = 0; k < dataABC.Rows.Count; k++)
                    {
                        oSheet.Cells[k + 2, a] = dataABC.Rows[k][s];
                        //oSheet.Cells[k + 1, s + 1].Borders.LineStyle = Excel.Constants.xlSolid;

                        if (k != dataABC.Rows.Count) oSheet.Cells[k + 2, 1] = k + 1;
                    }
                    oSheet.Columns.AutoFit();
                    oSheet.Rows.AutoFit();
                }

                oXL.DisplayAlerts = false;
                oWB.Save();
                //  oSheet.PrintOut();
                //  oWB.Close(SaveChanges: false);
                //  oXL.Application.Quit();
            }
            catch (System.Exception excep) 
            { 
                MessageBox.Show(excep.Message + " Məlumatlar Excel-ə doldurularkən, Excel-ə toxunmaq olmaz !!! "); 
            }
        }

        private void lBLNBankLayihələrToolStripMenuItem_Click(object sender, EventArgs e)
        {

            String name = "Кредитный портфель";
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            "1.xlsx" +
                            ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(constr);
            OleDbCommand oconn = new OleDbCommand("Select [ID клиента], [Наименование клиента], [Дата заключения контракта], [Дата окончания контракта], [Сумма контракта в манатном эквиваленте], [Остаток основного долга на дату в манатном эквиваленте], [Остаток просрочки на дату в манатном эквиваленте], [Начисленные проценты в манатном эквиваленте], [Просроченные проценты в манатном эквиваленте], [Штрафные проценты в манатном эквиваленте], [Срок просрочки процентов], [Последняя дата погашения процентов (или просроченных процентов)], [Последняя дата погашения основного долга], [Статус проблемы], [Проблема# Куратор] From [" + name + "$] WHERE [Способ выдачи кредита] like '%LEASING%'", con);
            con.Open();

            OleDbDataAdapter abc = new OleDbDataAdapter(oconn);
            DataTable dataABC = new DataTable();
            dataABC.Clear();
            try { abc.Fill(dataABC); }
            catch { MessageBox.Show("1.xlsx faylda 'Последняя дата погашения основного долга' xanasına bax"); }
            //dataGridView1.DataSource = dataABC;
            ///////////////////datatableni sortlamaq ucun
            dataABC.DefaultView.Sort = "[Наименование клиента] ASC";
            dataABC = dataABC.DefaultView.ToTable(true);
            dataGridView1.DataSource = dataABC;
            con.Close();

            try
            {
                int s = 0, k = 0, a = 0;

                try { File.Copy("Portfel.xlsm", "C:\\Users\\" + Environment.UserName + "\\Desktop\\PORTFEL (LB+LN).xlsm", true); }
                catch { MessageBox.Show("'Portfel.xlsm' tapılmadı."); }

                //Get a new workbook.
                oXL = new Excel.Application();
                oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\PORTFEL (LB+LN).xlsm"));
                oSheet = (Excel._Worksheet)oWB.Sheets[1];

                oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                oXL.Visible = true;
                oSheet.Activate();
                oSheet.Range["A1"].Select();
                oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;

                //oSheet.Cells.EntireColumn.NumberFormat = "@";

                for (s = 0; s < dataABC.Columns.Count; s++)
                {
                    a = s + 2;

                    oSheet.Cells[1, a] = dataABC.Columns[s].ColumnName;
                    oSheet.Cells[1, 1] = "Nö";

                    for (k = 0; k < dataABC.Rows.Count; k++)
                    {
                        oSheet.Cells[k + 2, a] =dataABC.Rows[k][s];
                        if (k != dataABC.Rows.Count) { oSheet.Cells[k + 2, 1] = k + 1; }

                        //oSheet.Cells[k + 2, a].Borders.LineStyle = Excel.Constants.xlSolid;
                    }
                }

                double cem1 = 0, cem2 = 0, cem3 = 0, cem4 = 0, cem5 = 0;
                for (k = 0; k < dataABC.Rows.Count; k++)
                {
                    try { cem1 += Convert.ToDouble(dataABC.Rows[k]["Остаток основного долга на дату в манатном эквиваленте"]); }
                    catch { }
                    try {cem2 += Convert.ToDouble(dataABC.Rows[k]["Остаток просрочки на дату в манатном эквиваленте"]);}
                    catch { }
                    try {cem3 += Convert.ToDouble(dataABC.Rows[k]["Начисленные проценты в манатном эквиваленте"]);}
                    catch { }
                    try {cem4 += Convert.ToDouble(dataABC.Rows[k]["Просроченные проценты в манатном эквиваленте"]);}
                    catch { }
                    try { cem5 += Convert.ToDouble(dataABC.Rows[k]["Штрафные проценты в манатном эквиваленте"]); }
                    catch { }
                }

                oSheet.Cells[dataABC.Rows.Count + 2, 7] = cem1.ToString();
                oSheet.Cells[dataABC.Rows.Count + 2, 8] = cem2.ToString();
                oSheet.Cells[dataABC.Rows.Count + 2, 9] = cem3.ToString();
                oSheet.Cells[dataABC.Rows.Count + 2, 10] = cem4.ToString();
                oSheet.Cells[dataABC.Rows.Count + 2, 11] = cem5.ToString();
                oSheet.Cells[dataABC.Rows.Count + 2, 12] = (cem1 + cem2 + cem3 + cem4 + cem5).ToString();
                oSheet.Range[oSheet.Cells[dataABC.Rows.Count + 2, 12], oSheet.Cells[dataABC.Rows.Count + 2, 14]].Merge();

                oSheet.Cells[dataABC.Rows.Count + 2, 7].Font.Bold = true;
                oSheet.Cells[dataABC.Rows.Count + 2, 8].Font.Bold = true;
                oSheet.Cells[dataABC.Rows.Count + 2, 9].Font.Bold = true;
                oSheet.Cells[dataABC.Rows.Count + 2, 10].Font.Bold = true;
                oSheet.Cells[dataABC.Rows.Count + 2, 11].Font.Bold = true;
                oSheet.Cells[dataABC.Rows.Count + 2, 12].Font.Bold = true;
                oSheet.Columns.AutoFit();
                oSheet.Rows.AutoFit();

                oXL.DisplayAlerts = false;
                oWB.Save();
                //  oSheet.PrintOut();
                //  oWB.Close(SaveChanges: false);
                //  oXL.Application.Quit();
            }
            catch (System.Exception excep)
            {
                MessageBox.Show(excep.Message + Environment.NewLine + " Məlumatlar Excel-ə doldurularkən, Excel-ə toxunmaq olmaz !!! ");
            }
        }

        private void gecikmişLizinqLayihələrToolStripMenuItem_Click(object sender, EventArgs e)
        {
            gecikmede_olanlar();
        }

        private void mehkemedeOlanlarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            progressBar1.Value = 0;

            
            
            
            MyData.selectCommand("baza.accdb", "SELECT * FROM PortfelStatus WHERE Status Like '%M%' order by Layihe desc");
            MyData.dtmain= new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmain);
            //MyData.dtmain = MyData.dtmain.DefaultView.ToTable(true);

            String name = "licschkre";
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            "2.xlsx" +
                            ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(constr);
            DataTable dataABC = new DataTable();
            dataABC.Clear();
            for (int i = 0; i < MyData.dtmain.Rows.Count; i++)
            {
                OleDbCommand oconn = new OleDbCommand("Select [Ö#G#], [Layihe], Adı, [Lizinqin məbləği], [Qalıq], [% məbləği], [V#K#Qalıq], [V#K#% məbləği], [Cərimə % məbləği], [Dəbbə məbləği], [Son əməl#tarixi], [Qrafikda olan məbləğ], [overdueday] From [" + name + "$] WHERE Layihe Like '%" + MyData.dtmain.Rows[i]["Layihe"].ToString() + "%'", con);

                con.Open();
                OleDbDataAdapter abc = new OleDbDataAdapter(oconn);

                abc.Fill(dataABC);
                
                ///////////////////datatableni sortlamaq ucun
                dataABC.DefaultView.Sort = "Layihe desc";
                dataABC = dataABC.DefaultView.ToTable(true);
                //dataGridView1.DataSource = dataABC;
                con.Close();

                try
                {
                    progressBar1.Value += 100 / MyData.dtmain.Rows.Count;
                }
                catch{}
            }

            progressBar1.Value = 100;

            try
            {
                int s, k, a = 0;

                try { File.Copy("Portfel.xlsm", "C:\\Users\\" + Environment.UserName + "\\Desktop\\PORTFEL (Mehkemede olanlar).xlsm", true); }
                catch { MessageBox.Show("'Portfel.xlsm' tapılmadı."); }

                //Get a new workbook.
                oXL = new Excel.Application();
                oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\PORTFEL (Mehkemede olanlar).xlsm"));
                oSheet = (Excel._Worksheet)oWB.Sheets[1];

                oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                oXL.Visible = true;
                oSheet.Activate();
                oSheet.Range["A1"].Select();
                oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;

                //oSheet.Cells.EntireColumn.NumberFormat = "@";

                for (s = 0; s < dataABC.Columns.Count; s++)
                {
                    a = s + 2;

                    oSheet.Cells[1, a] = dataABC.Columns[s].ColumnName;
                    oSheet.Cells[1, 1] = "Nö";

                    for (k = 0; k < dataABC.Rows.Count; k++)
                    {
                        oSheet.Cells[k + 2, a] = dataABC.Rows[k][s].ToString();
                        if (dataABC.Columns[s].ColumnName == "Son əməl#tarixi") oSheet.Cells[k + 2, a] = "'" + dataABC.Rows[k][s].ToString();
                        //oSheet.Cells[k + 1, s + 1].Borders.LineStyle = Excel.Constants.xlSolid;

                        if (k != dataABC.Rows.Count) oSheet.Cells[k + 2, 1] = k + 1;
                    }
                    oSheet.Columns.AutoFit();
                    oSheet.Rows.AutoFit();
                }

                oXL.DisplayAlerts = false;
                oWB.Save();
                //  oSheet.PrintOut();
                //  oWB.Close(SaveChanges: false);
                //  oXL.Application.Quit();        
            }
            catch (System.Exception excep)
            {
                MessageBox.Show(excep.Message + " Məlumatlar Excel-ə doldurularkən, Excel-ə toxunmaq olmaz !!! ");
            }

        }

        private void lSProblemliyəVerilənLayihələrToolStripMenuItem_Click(object sender, EventArgs e)
        {
            progressBar1.Value = 0;

            
            
            
            MyData.selectCommand("baza.accdb", "SELECT * FROM PortfelStatus WHERE Status Like '%P%' order by Layihe desc");
            MyData.dtmain = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmain);
            //MyData.dtmain = MyData.dtmain.DefaultView.ToTable(true);

            String name = "licschkre";
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            "2.xlsx" +
                            ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(constr);
            DataTable dataABC = new DataTable();
            dataABC.Clear();
            for (int i = 0; i < MyData.dtmain.Rows.Count; i++)
            {
                OleDbCommand oconn = new OleDbCommand("Select [Ö#G#], [Layihe], Adı, [Lizinqin məbləği], [Qalıq], [% məbləği], [V#K#Qalıq], [V#K#% məbləği], [Cərimə % məbləği], [Dəbbə məbləği], [Son əməl#tarixi], [Qrafikda olan məbləğ], [overdueday] From [" + name + "$] WHERE Layihe Like '%" + MyData.dtmain.Rows[i]["Layihe"].ToString() + "%'", con);

                con.Open();
                OleDbDataAdapter abc = new OleDbDataAdapter(oconn);

                abc.Fill(dataABC);

                ///////////////////datatableni sortlamaq ucun
                dataABC.DefaultView.Sort = "Layihe desc";
                dataABC = dataABC.DefaultView.ToTable(true);
                //dataGridView1.DataSource = data;
                con.Close();

                try
                {
                    progressBar1.Value += 100 / MyData.dtmain.Rows.Count;
                }
                catch { }
            }

            progressBar1.Value = 100;

            try
            {
                int s, k, a = 0;

                try { File.Copy("Portfel.xlsm", "C:\\Users\\" + Environment.UserName + "\\Desktop\\PORTFEL (Problemli).xlsm", true); }
                catch { MessageBox.Show("'Portfel.xlsm' tapılmadı."); }

                //Get a new workbook.
                oXL = new Excel.Application();
                oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\PORTFEL (Problemli).xlsm"));
                oSheet = (Excel._Worksheet)oWB.Sheets[1];

                oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                oXL.Visible = true;
                oSheet.Activate();
                oSheet.Range["A1"].Select();
                oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;

                //oSheet.Cells.EntireColumn.NumberFormat = "@";

                for (s = 0; s < dataABC.Columns.Count; s++)
                {
                    a = s + 2;

                    oSheet.Cells[1, a] = dataABC.Columns[s].ColumnName;
                    oSheet.Cells[1, 1] = "Nö";

                    for (k = 0; k < dataABC.Rows.Count; k++)
                    {
                        oSheet.Cells[k + 2, a] = dataABC.Rows[k][s].ToString();
                        if (dataABC.Columns[s].ColumnName == "Son əməl#tarixi") oSheet.Cells[k + 2, a] = "'" + dataABC.Rows[k][s].ToString();
                        //oSheet.Cells[k + 1, s + 1].Borders.LineStyle = Excel.Constants.xlSolid;

                        if (k != dataABC.Rows.Count) oSheet.Cells[k + 2, 1] = k + 1;
                    }
                    oSheet.Columns.AutoFit();
                    oSheet.Rows.AutoFit();
                }

                oXL.DisplayAlerts = false;
                oWB.Save();
                //  oSheet.PrintOut();
                //  oWB.Close(SaveChanges: false);
                //  oXL.Application.Quit();        
            }
            catch (System.Exception excep)
            {
                MessageBox.Show(excep.Message + " Məlumatlar Excel-ə doldurularkən, Excel-ə toxunmaq olmaz !!! ");
            }

        }

        private void lSBizdəOlanLayihələrToolStripMenuItem_Click(object sender, EventArgs e)
        {
            progressBar1.Value = 0;

            
            
            
            MyData.selectCommand("baza.accdb", "SELECT * FROM PortfelStatus WHERE Status Like '%L%' order by Layihe desc");
            MyData.dtmain= new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmain);
            //MyData.dtmain = MyData.dtmain.DefaultView.ToTable(true);

            String name = "licschkre";
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            "2.xlsx" +
                            ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(constr);
            DataTable dataABC = new DataTable();
            dataABC.Clear();
            for (int i = 0; i < MyData.dtmain.Rows.Count; i++)
            {
                OleDbCommand oconn = new OleDbCommand("Select [Ö#G#], [Layihe], Adı, [Lizinqin məbləği], [Qalıq], [% məbləği], [V#K#Qalıq], [V#K#% məbləği], [Cərimə % məbləği], [Dəbbə məbləği], [Son əməl#tarixi], [Qrafikda olan məbləğ], [overdueday] From [" + name + "$] WHERE Layihe Like '%" + MyData.dtmain.Rows[i]["Layihe"].ToString() + "%'", con);

                con.Open();
                OleDbDataAdapter abc = new OleDbDataAdapter(oconn);

                abc.Fill(dataABC);

                ///////////////////datatableni sortlamaq ucun
                dataABC.DefaultView.Sort = "Layihe desc";
                dataABC = dataABC.DefaultView.ToTable(true);
                //dataGridView1.DataSource = data;
                con.Close();

                try
                {
                    progressBar1.Value = progressBar1.Value + 100 / MyData.dtmain.Rows.Count;
                }
                catch { }
            }

            progressBar1.Value = 100;

            try
            {
                int s, k, a = 0;

                try { File.Copy("Portfel.xlsm", "C:\\Users\\" + Environment.UserName + "\\Desktop\\PORTFEL (bizde olanlar).xlsm", true); }
                catch { MessageBox.Show("'Portfel.xlsm' tapılmadı."); }

                //Get a new workbook.
                oXL = new Excel.Application();
                oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\PORTFEL (bizde olanlar).xlsm"));
                oSheet = (Excel._Worksheet)oWB.Sheets[1];

                oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                oXL.Visible = true;
                oSheet.Activate();
                oSheet.Range["A1"].Select();
                oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;

                //oSheet.Cells.EntireColumn.NumberFormat = "@";

                for (s = 0; s < dataABC.Columns.Count; s++)
                {
                    a = s + 2;

                    oSheet.Cells[1, a] = dataABC.Columns[s].ColumnName;
                    oSheet.Cells[1, 1] = "Nö";

                    for (k = 0; k < dataABC.Rows.Count; k++)
                    {
                        oSheet.Cells[k + 2, a] = dataABC.Rows[k][s].ToString();
                        if (dataABC.Columns[s].ColumnName == "Son əməl#tarixi") oSheet.Cells[k + 2, a] = "'" + dataABC.Rows[k][s].ToString();
                        //oSheet.Cells[k + 1, s + 1].Borders.LineStyle = Excel.Constants.xlSolid;

                        if (k != dataABC.Rows.Count) oSheet.Cells[k + 2, 1] = k + 1;
                    }
                    oSheet.Columns.AutoFit();
                    oSheet.Rows.AutoFit();
                }

                oXL.DisplayAlerts = false;
                oWB.Save();
                //  oSheet.PrintOut();
                //  oWB.Close(SaveChanges: false);
                //  oXL.Application.Quit();        
            }
            catch (System.Exception excep)
            {
                MessageBox.Show(excep.Message + " Məlumatlar Excel-ə doldurularkən, Excel-ə toxunmaq olmaz !!! ");
            }

        }

        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            gecikmede_olanlar();
        }

        private void printToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string gecikme = "";
            DateTime dt = DateTime.Now;

            try { File.Copy("New Emphty.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Gecikme - " + dt.ToShortDateString() + ".doc", true); }
            catch { MessageBox.Show("'\\192.168.10.5\\Common\\AGLizinq\\New Emphty.doc' tapılmadı."); }

            object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\Gecikme - " + dt.ToShortDateString() + ".doc";

            string tarix = "", OdenisGunu = "";
            //try
            //{

            String name = "licschkre";
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            "2.xlsx" +
                            ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(constr);
            OleDbCommand oconn = new OleDbCommand("Select Adı, [V#K#Qalıq], [V#K#% məbləği], [Dəbbə məbləği], [Cərimə % məbləği], [Qrafikda olan məbləğ], [Verilmə tarixi], [overdueday], [Qalıq], [% məbləği], Layihe, [K#p#b#tarixi], [Son əməl#tarixi], [Uzadılma tarixi] From [" + name + "$] WHERE NOT Layihe Like '%S-01/15-048/15%' and NOT Layihe Like '%S-032/14%' and Layihe Like '%S-%' or Layihe Like '%A-%'  or Layihe like '%126/08%' or Layihe like '%106/08%' or Layihe like '%158%' or Layihe like '%079%' or Layihe like '%174%' or Layihe like '%105/08%' or Layihe like '%105/08%' or Layihe like '%198%' or Layihe like '%193%' or Layihe like '%161%' or Layihe like '%196%' or Layihe like '%114%' or Layihe like '%L-01/16%' or Layihe like '%055%' or Layihe like '%01/13-167/10%'or Layihe like '%02/14-167/10%'or Layihe like '%03/14-167/10%'or Layihe like '%04/16-167/10%'", con);
            con.Open();
            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
            DataTable data1 = new DataTable();
            data1.Clear();
            sda.Fill(data1);
            con.Close();

            int k = 0, k2 = 0;
            double CemiGecikme = 0, Gecikme100 = 0, CemiQaliq = 0, NisbetFaiz = 0;


            for (int i = 0; i < data1.Rows.Count; i++)
            {
                if (data1.Rows[i]["Uzadılma tarixi"].ToString() != "") OdenisGunu = data1.Rows[i]["Uzadılma tarixi"].ToString().Substring(0, 2);
                else OdenisGunu = data1.Rows[i]["K#p#b#tarixi"].ToString().Substring(0, 2);

                CemiQaliq += Math.Round(Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"]) + Convert.ToDouble(data1.Rows[i]["Qalıq"]) + Convert.ToDouble(data1.Rows[i]["% məbləği"]), 2);

                if ((Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"]) > 0))
                {
                    ////ABDULLAYEV TOFIQ UCUN XUSUSI YAZILIM GECIKMEDE FAIZ BORCUN GORUNMEMESI UCUN.........................
                    if (data1.Rows[i]["Layihe"].ToString() == "'S-008/14") { gecikme += Environment.NewLine + OdenisGunu + " ♦ " + data1.Rows[i]["Adı"].ToString() + " - " + Math.Round((Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"])), 2).ToString() + "/" + Math.Round((Convert.ToDouble(data1.Rows[i]["% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"])), 2).ToString() + " ♦ aylıq - " + data1.Rows[i]["Qrafikda olan məbləğ"].ToString() + " ♦ " + data1.Rows[i]["Layihe"].ToString().Substring(1, data1.Rows[i]["Layihe"].ToString().Length - 1) + " ♦ son ödəmə - " + data1.Rows[i]["Son əməl#tarixi"].ToString(); }
                    else { gecikme += Environment.NewLine + OdenisGunu + " ♦ " + data1.Rows[i]["Adı"].ToString() + " - " + Math.Round((Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"])), 2).ToString() + "/" + Math.Round((Convert.ToDouble(data1.Rows[i]["% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"])), 2).ToString() + " ♦ aylıq - " + data1.Rows[i]["Qrafikda olan məbləğ"].ToString() + " ♦ " + data1.Rows[i]["Layihe"].ToString().Substring(1, data1.Rows[i]["Layihe"].ToString().Length - 1) + " ♦ son ödəmə - " + data1.Rows[i]["Son əməl#tarixi"].ToString(); }
                    
                    k += 1;
                    CemiGecikme += Math.Round(Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"]), 2);
                }

                if (Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"]) > 0 && Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"]) < 100)
                {
                    k2 += 1;
                    Gecikme100 += Math.Round(Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"]), 2);
                }
            }
            //notifyIcon1.Icon = SystemIcons.;
            NisbetFaiz = Math.Round(CemiGecikme / CemiQaliq * 100, 2);   //gecikmenin faizinin tapilmasi
            //pertfel vaxti ucun
            try
            {
                String name2 = "licschkre";
                String constr2 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                "2.xlsx" +
                                ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

                OleDbConnection con2 = new OleDbConnection(constr2);
                OleDbCommand oconn2 = new OleDbCommand("Select overdueday From [" + name2 + "$]", con2);
                con2.Open();
                OleDbDataAdapter sda2 = new OleDbDataAdapter(oconn2);
                DataTable data2 = new DataTable();
                data2.Clear();
                sda2.Fill(data2);
                con2.Close();
                tarix = "PORTFEL TARİXİ ♦ " + data2.Rows[0]["overdueday"].ToString();
            }
            catch { }

            string qeydler = "";
            
            
            
            MyData.selectCommand("baza.accdb", "Select * from Qeydler");
            MyData.dtmain= new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmain);

            if (MyData.dtmain.Rows.Count != 0)
            {
                for (int i = 0; i < MyData.dtmain.Rows.Count; i++)
                {
                    qeydler += Environment.NewLine + MyData.dtmain.Rows[i][1].ToString().Substring(0, 10) + " - " + MyData.dtmain.Rows[i][2].ToString();
                }

                qeydler = "QEYDLƏR:" + qeydler;
            }

            Microsoft.Office.Interop.Word._Application oWord;
            object oMissing = Type.Missing;
            oWord = new Microsoft.Office.Interop.Word.Application();
            oWord.Visible = false;
            oWord.Documents.Open(FileName);
            oWord.Selection.TypeText(tarix + Environment.NewLine + gecikme + Environment.NewLine + Environment.NewLine + "İNFO:" + Environment.NewLine + "Real müştərilər - " + data1.Rows.Count.ToString() + " ədəd, məbləğ - " + CemiQaliq + " azn" + Environment.NewLine + "Gecikmədə olanlar - " + k.ToString() + " ədəd, məbləğ - " + CemiGecikme + " azn  (Gecikmə faizlə - " + NisbetFaiz.ToString() + "%)" + Environment.NewLine + "O cümlədən, 100 AZN-ə kimi gecikənlər - " + k2.ToString() + " ədəd, məbləğ - " + Gecikme100 + " azn" + Environment.NewLine + Environment.NewLine + qeydler);
            oWord.PrintOut();
            oWord.ActiveDocument.Save();
            oWord.Quit();
            //MessageBox.Show("The text is inserted.");

            //}
            //catch { MessageBox.Show("Səhv var."); }
        }

        private void saglamToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saglam_olanlar();
        }

        private void lSWordToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string gecikme = "";
            DateTime dt = DateTime.Now;

            try { File.Copy("New Emphty.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Gecikme - " + dt.ToShortDateString() + ".doc", true); }
            catch { MessageBox.Show("'\\192.168.10.5\\Common\\AGLizinq\\New Emphty.doc' tapılmadı."); }

            object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\Gecikme - " + dt.ToShortDateString() + ".doc";

            string tarix = "", OdenisGunu = "";
            //try
            //{

            String name = "licschkre";
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            "2.xlsx" +
                            ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(constr);
            OleDbCommand oconn = new OleDbCommand("Select Adı, [V#K#Qalıq], [V#K#% məbləği], [Dəbbə məbləği], [Cərimə % məbləği], [Qrafikda olan məbləğ], [Verilmə tarixi], [overdueday], [Qalıq], [% məbləği], Layihe, [K#p#b#tarixi], [Son əməl#tarixi], [Uzadılma tarixi] From [" + name + "$]", con);
            con.Open();
            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
            DataTable data1 = new DataTable();
            data1.Clear();
            sda.Fill(data1);
            con.Close();

            int k = 0, k2 = 0;
            double CemiGecikme = 0, Gecikme100 = 0, CemiQaliq = 0, NisbetFaiz = 0;


            for (int i = 0; i < data1.Rows.Count; i++)
            {
                if (data1.Rows[i]["Uzadılma tarixi"].ToString() != "") OdenisGunu = data1.Rows[i]["Uzadılma tarixi"].ToString().Substring(0, 2);
                else OdenisGunu = data1.Rows[i]["K#p#b#tarixi"].ToString().Substring(0, 2);

                CemiQaliq += Math.Round(Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"]) + Convert.ToDouble(data1.Rows[i]["Qalıq"]) + Convert.ToDouble(data1.Rows[i]["% məbləği"]), 2);

                if ((Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"]) > 0))
                {
                    ////ABDULLAYEV TOFIQ UCUN XUSUSI YAZILIM GECIKMEDE FAIZ BORCUN GORUNMEMESI UCUN.........................
                    if (data1.Rows[i]["Layihe"].ToString() == "'S-008/14") { gecikme += Environment.NewLine + OdenisGunu + " ♦ " + data1.Rows[i]["Adı"].ToString() + " - " + Math.Round((Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"])), 2).ToString() + "/" + Math.Round((Convert.ToDouble(data1.Rows[i]["% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"])), 2).ToString() + " ♦ aylıq - " + data1.Rows[i]["Qrafikda olan məbləğ"].ToString() + " ♦ " + data1.Rows[i]["Layihe"].ToString().Substring(1, data1.Rows[i]["Layihe"].ToString().Length - 1) + " ♦ son ödəmə - " + data1.Rows[i]["Son əməl#tarixi"].ToString(); }
                    else { gecikme += Environment.NewLine + OdenisGunu + " ♦ " + data1.Rows[i]["Adı"].ToString() + " - " + Math.Round((Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"])), 2).ToString() + "/" + Math.Round((Convert.ToDouble(data1.Rows[i]["% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"])), 2).ToString() + " ♦ aylıq - " + data1.Rows[i]["Qrafikda olan məbləğ"].ToString() + " ♦ " + data1.Rows[i]["Layihe"].ToString().Substring(1, data1.Rows[i]["Layihe"].ToString().Length - 1) + " ♦ son ödəmə - " + data1.Rows[i]["Son əməl#tarixi"].ToString(); }
                    
                    k += 1;
                    CemiGecikme += Math.Round(Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"]), 2);
                }

                if (Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"]) > 0 && Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"]) < 100)
                {
                    k2 += 1;
                    Gecikme100 += Math.Round(Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"]), 2);
                }
            }
            //notifyIcon1.Icon = SystemIcons.;
            NisbetFaiz = Math.Round(CemiGecikme / CemiQaliq * 100, 2);   //gecikmenin faizinin tapilmasi
            //pertfel vaxti ucun
            try
            {
                String name2 = "licschkre";
                String constr2 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                "2.xlsx" +
                                ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

                OleDbConnection con2 = new OleDbConnection(constr2);
                OleDbCommand oconn2 = new OleDbCommand("Select overdueday From [" + name2 + "$]", con2);
                con2.Open();
                OleDbDataAdapter sda2 = new OleDbDataAdapter(oconn2);
                DataTable data2 = new DataTable();
                data2.Clear();
                sda2.Fill(data2);
                con2.Close();
                tarix = "PORTFEL TARİXİ ♦ " + data2.Rows[0]["overdueday"].ToString();
            }
            catch { }

            string qeydler = "";
            
            
            
            MyData.selectCommand("baza.accdb", "Select * from Qeydler");
            MyData.dtmain= new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmain);

            if (MyData.dtmain.Rows.Count != 0)
            {
                for (int i = 0; i < MyData.dtmain.Rows.Count; i++)
                {
                    qeydler += Environment.NewLine + MyData.dtmain.Rows[i][1].ToString().Substring(0, 10) + " - " + MyData.dtmain.Rows[i][2].ToString();
                }

                qeydler = "QEYDLƏR:" + qeydler;
            }

            Microsoft.Office.Interop.Word._Application oWord;
            object oMissing = Type.Missing;
            oWord = new Microsoft.Office.Interop.Word.Application();
            oWord.Visible = true;
            oWord.Documents.Open(FileName);
            oWord.Selection.TypeText(tarix + Environment.NewLine + gecikme + Environment.NewLine + Environment.NewLine + "İNFO:" + Environment.NewLine + "Ümumi müştərilər - " + data1.Rows.Count.ToString() + " ədəd, məbləğ - " + CemiQaliq + " azn" + Environment.NewLine + "Gecikmədə olanlar - " + k.ToString() + " ədəd, məbləğ - " + CemiGecikme + " azn  (Gecikmə faizlə - " + NisbetFaiz.ToString() + "%)" + Environment.NewLine + "O cümlədən, 100 AZN-ə kimi gecikənlər - " + k2.ToString() + " ədəd, məbləğ - " + Gecikme100 + " azn" + Environment.NewLine + Environment.NewLine + qeydler);
            oWord.ActiveDocument.Save();
            //oWord.Quit();
            //MessageBox.Show("The text is inserted.");

            //}
            //catch { MessageBox.Show("Səhv var."); }

        }

        private void digərToolStripMenuItem_Click(object sender, EventArgs e)
        {
            radioButton4.Checked = true;
            panel1.Visible = true;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (radioButton4.Checked == true)
            {
                try
                {
                String name = "Кредитный портфель";
                String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                "1.xlsx" +
                                ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

                OleDbConnection con = new OleDbConnection(constr);
                OleDbCommand oconn = new OleDbCommand("Select [ID клиента], [Наименование клиента], [Дата заключения контракта], [Дата окончания контракта], [Валюта контракта], [Сумма контракта], [Сумма контракта в манатном эквиваленте], [Остаток основного долга на дату], [Остаток основного долга на дату в манатном эквиваленте], [Остаток просрочки на дату], [Остаток просрочки на дату в манатном эквиваленте], [Начисленные проценты], [Начисленные проценты в манатном эквиваленте], [Просроченные проценты], [Просроченные проценты в манатном эквиваленте], [Штрафные проценты], [Штрафные проценты в манатном эквиваленте], [Срок просрочки процентов], [Последняя дата погашения процентов (или просроченных процентов)], [Куратор кредита], [Способ выдачи кредита] From [" + name + "$] WHERE [" + comboBox1.Text + "] like '%" + textBox2.Text + "%' and [" + comboBox2.Text + "] like '%" + textBox3.Text + "%' and [" + comboBox3.Text + "] like '%" + textBox4.Text + "%'", con);
                con.Open();

                OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                DataTable dataABC = new DataTable();
                dataABC.Clear();
                sda.Fill(dataABC);
                dataGridView1.DataSource = dataABC;
                con.Close();

                }
                catch { MessageBox.Show("AGBANK: Yuxarıdan üç xana boş olmamalıdır..."); }
            }

            else
            {
                try
                {
                    String name = "licschkre";
                    String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    "2.xlsx" +
                                    ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

                    OleDbConnection con = new OleDbConnection(constr);
                    OleDbCommand oconn = new OleDbCommand("Select Adı, Layihe, [Verilmə tarixi], [K#p#b#tarixi], [Qrafikda olan məbləğ], [Son əməl#tarixi], [Uzadılma tarixi], [Lizinqin məbləği], [Qalıq], [V#K#Qalıq], [% məbləği], [V#K#% məbləği], [Cərimə % məbləği], [Dəbbə məbləği], [Lizinqin növü], [Lizinq obyektin dəyəri], [Val#], [overdueday], [L#K#], [Ö#G#] From [" + name + "$] WHERE [" + comboBox1.Text + "] like '%" + textBox2.Text + "%' and [" + comboBox2.Text + "] like '%" + textBox3.Text + "%' and [" + comboBox3.Text + "] like '%" + textBox4.Text + "%'", con);
                    con.Open();

                    OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                    DataTable dataABC = new DataTable();
                    dataABC.Clear();
                    sda.Fill(dataABC);
                    dataGridView1.DataSource = dataABC;
                    con.Close();

                }
                catch { MessageBox.Show("AGLIZINQ: Yuxarıdan üç xana boş olmamalıdır..."); }
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";

            panel1.Visible = false;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            if (radioButton4.Checked == true)
            {
                try
                {
                    int s = 0, k = 0, a = 0;

                    try { File.Copy("Portfel.xlsm", "C:\\Users\\" + Environment.UserName + "\\Desktop\\PORTFEL BANK Diger.xlsm", true); }
                    catch { MessageBox.Show("'Portfel.xlsm' tapılmadı."); }

                    //Get a new workbook.
                    oXL = new Excel.Application();
                    oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\PORTFEL BANK Diger.xlsm"));
                    oSheet = (Excel._Worksheet)oWB.Sheets[1];

                    oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                    oXL.Visible = true;
                    oSheet.Activate();
                    oSheet.Range["A1"].Select();
                    oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;

                    //oSheet.Cells.EntireColumn.NumberFormat = "@";

                    for (s = 0; s < dataGridView1.Columns.Count; s++)
                    {
                        a = s + 2;

                        oSheet.Cells[1, a] = dataGridView1.Columns[s].HeaderText;
                        oSheet.Cells[1, 1] = "Nö";

                        for (k = 0; k < dataGridView1.Rows.Count; k++)
                        {
                            oSheet.Cells[k + 2, a] = dataGridView1.Rows[k].Cells[s].Value;
                            if (k != dataGridView1.Rows.Count) { oSheet.Cells[k + 2, 1] = k + 1; }

                            //oSheet.Cells[k + 2, a].Borders.LineStyle = Excel.Constants.xlSolid;
                        }
                    }

                    double cem1 = 0, cem2 = 0, cem3 = 0, cem4 = 0, cem5 = 0;
                    for (k = 0; k < dataGridView1.Rows.Count; k++)
                    {
                        try { cem1 += Convert.ToDouble(dataGridView1.Rows[k].Cells["Остаток основного долга на дату в манатном эквиваленте"].Value); }
                        catch { }
                        try { cem2 += Convert.ToDouble(dataGridView1.Rows[k].Cells["Остаток просрочки на дату в манатном эквиваленте"].Value); }
                        catch { }
                        try { cem3 += Convert.ToDouble(dataGridView1.Rows[k].Cells["Начисленные проценты в манатном эквиваленте"].Value); }
                        catch { }
                        try { cem4 += Convert.ToDouble(dataGridView1.Rows[k].Cells["Просроченные проценты в манатном эквиваленте"].Value); }
                        catch { }
                        try { cem5 += Convert.ToDouble(dataGridView1.Rows[k].Cells["Штрафные проценты в манатном эквиваленте"].Value); }
                        catch { }
                    }

                    oSheet.Cells[dataGridView1.Rows.Count + 2, 10] = cem1.ToString();
                    oSheet.Cells[dataGridView1.Rows.Count + 2, 11] = cem2.ToString();
                    oSheet.Cells[dataGridView1.Rows.Count + 2, 12] = cem3.ToString();
                    oSheet.Cells[dataGridView1.Rows.Count + 2, 13] = cem4.ToString();
                    oSheet.Cells[dataGridView1.Rows.Count + 2, 14] = cem5.ToString();
                    oSheet.Cells[dataGridView1.Rows.Count + 2, 15] = (cem1 + cem2 + cem3 + cem4 + cem5).ToString();
                    oSheet.Range[oSheet.Cells[dataGridView1.Rows.Count + 2, 15], oSheet.Cells[dataGridView1.Rows.Count + 2, 21]].Merge();

                    oSheet.Cells[dataGridView1.Rows.Count + 2, 10].Font.Bold = true;
                    oSheet.Cells[dataGridView1.Rows.Count + 2, 11].Font.Bold = true;
                    oSheet.Cells[dataGridView1.Rows.Count + 2, 12].Font.Bold = true;
                    oSheet.Cells[dataGridView1.Rows.Count + 2, 13].Font.Bold = true;
                    oSheet.Cells[dataGridView1.Rows.Count + 2, 14].Font.Bold = true;
                    oSheet.Cells[dataGridView1.Rows.Count + 2, 15].Font.Bold = true;
                    oSheet.Columns.AutoFit();
                    oSheet.Rows.AutoFit();

                    oXL.DisplayAlerts = false;
                    oWB.Save();
                    //  oSheet.PrintOut();
                    //  oWB.Close(SaveChanges: false);
                    //  oXL.Application.Quit();
                }
                catch (System.Exception excep)
                {
                    MessageBox.Show(excep.Message + Environment.NewLine + "AGBANK Məlumatlar Excel-ə doldurularkən, Excel-ə toxunmaq olmaz !!! ");
                }
            }

            if (radioButton5.Checked == true)
            {
                try
                {
                    int s = 0, k = 0, a = 0;

                    try { File.Copy("Portfel.xlsm", "C:\\Users\\" + Environment.UserName + "\\Desktop\\PORTFEL Lizinq Diger.xlsm", true); }
                    catch { MessageBox.Show("'Portfel.xlsm' tapılmadı."); }

                    //Get a new workbook.
                    oXL = new Excel.Application();
                    oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\PORTFEL Lizinq Diger.xlsm"));
                    oSheet = (Excel._Worksheet)oWB.Sheets[1];

                    oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                    oXL.Visible = true;
                    oSheet.Activate();
                    oSheet.Range["A1"].Select();
                    oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;

                    //oSheet.Cells.EntireColumn.NumberFormat = "@";

                    for (s = 0; s < dataGridView1.Columns.Count; s++)
                    {
                        a = s + 2;

                        oSheet.Cells[1, a] = dataGridView1.Columns[s].HeaderText;
                        oSheet.Cells[1, 1] = "Nö";

                        for (k = 0; k < dataGridView1.Rows.Count; k++)
                        {
                            oSheet.Cells[k + 2, a] = dataGridView1.Rows[k].Cells[s].Value;
                            if (k != dataGridView1.Rows.Count) { oSheet.Cells[k + 2, 1] = k + 1; }

                            //oSheet.Cells[k + 2, a].Borders.LineStyle = Excel.Constants.xlSolid;
                        }
                    }

                    double cem1 = 0, cem2 = 0, cem3 = 0, cem4 = 0, cem5 = 0;
                    for (k = 0; k < dataGridView1.Rows.Count; k++)
                    {
                        try { cem1 += Convert.ToDouble(dataGridView1.Rows[k].Cells["Qalıq"].Value); }
                        catch { }
                        try { cem2 += Convert.ToDouble(dataGridView1.Rows[k].Cells["V#K#Qalıq"].Value); }
                        catch { }
                        try { cem3 += Convert.ToDouble(dataGridView1.Rows[k].Cells["% məbləği"].Value); }
                        catch { }
                        try { cem4 += Convert.ToDouble(dataGridView1.Rows[k].Cells["V#K#% məbləği"].Value); }
                        catch { }
                        try { cem5 += Convert.ToDouble(dataGridView1.Rows[k].Cells["Cərimə % məbləği"].Value) + Convert.ToDouble(dataGridView1.Rows[k].Cells["Dəbbə məbləği"].Value); }
                        catch { }
                    }

                    oSheet.Cells[dataGridView1.Rows.Count + 2, 10] = cem1.ToString();
                    oSheet.Cells[dataGridView1.Rows.Count + 2, 11] = cem2.ToString();
                    oSheet.Cells[dataGridView1.Rows.Count + 2, 12] = cem3.ToString();
                    oSheet.Cells[dataGridView1.Rows.Count + 2, 13] = cem4.ToString();
                    oSheet.Cells[dataGridView1.Rows.Count + 2, 14] = cem5.ToString();
                    oSheet.Cells[dataGridView1.Rows.Count + 2, 15] = (cem1 + cem2 + cem3 + cem4 + cem5).ToString();
                    oSheet.Range[oSheet.Cells[dataGridView1.Rows.Count + 2, 15], oSheet.Cells[dataGridView1.Rows.Count + 2, 21]].Merge();

                    oSheet.Cells[dataGridView1.Rows.Count + 2, 10].Font.Bold = true;
                    oSheet.Cells[dataGridView1.Rows.Count + 2, 11].Font.Bold = true;
                    oSheet.Cells[dataGridView1.Rows.Count + 2, 12].Font.Bold = true;
                    oSheet.Cells[dataGridView1.Rows.Count + 2, 13].Font.Bold = true;
                    oSheet.Cells[dataGridView1.Rows.Count + 2, 14].Font.Bold = true;
                    oSheet.Cells[dataGridView1.Rows.Count + 2, 15].Font.Bold = true;
                    oSheet.Columns.AutoFit();
                    oSheet.Rows.AutoFit();

                    oXL.DisplayAlerts = false;
                    oWB.Save();
                    //  oSheet.PrintOut();
                    //  oWB.Close(SaveChanges: false);
                    //  oXL.Application.Quit();
                }
                catch (System.Exception excep)
                {
                    MessageBox.Show(excep.Message + Environment.NewLine + "AGLIZINQ Məlumatlar Excel-ə doldurularkən, Excel-ə toxunmaq olmaz !!! ");
                }
            }
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton4.Checked == true)
            {
                String name = "Кредитный портфель";
                String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                "1.xlsx" +
                                ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

                OleDbConnection con = new OleDbConnection(constr);
                OleDbCommand oconn = new OleDbCommand("Select * From [" + name + "$]", con);
                con.Open();

                OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                DataTable dataABC = new DataTable();
                dataABC.Clear();
                sda.Fill(dataABC);

                comboBox1.Items.Clear();
                comboBox2.Items.Clear();
                comboBox3.Items.Clear();
                comboBox1.Text = dataABC.Columns[63].ColumnName;
                comboBox2.Text = dataABC.Columns[15].ColumnName;
                comboBox3.Text = dataABC.Columns[21].ColumnName;
                for (int i = 0; i < dataABC.Columns.Count; i++)
                {
                    comboBox1.Items.Add(dataABC.Columns[i].ColumnName);
                    comboBox2.Items.Add(dataABC.Columns[i].ColumnName);
                    comboBox3.Items.Add(dataABC.Columns[i].ColumnName);
                }

            }

            else
            {
                String name = "licschkre";
                String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                "2.xlsx" +
                                ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

                OleDbConnection con = new OleDbConnection(constr);
                OleDbCommand oconn = new OleDbCommand("Select * From [" + name + "$]", con);
                con.Open();

                OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                DataTable dataABC = new DataTable();
                dataABC.Clear();
                sda.Fill(dataABC);

                comboBox1.Items.Clear();
                comboBox2.Items.Clear();
                comboBox3.Items.Clear();
                comboBox1.Text = dataABC.Columns[1].ColumnName;
                comboBox2.Text = dataABC.Columns[39].ColumnName;
                comboBox3.Text = dataABC.Columns[53].ColumnName;
                for (int i = 0; i < dataABC.Columns.Count; i++)
                {
                    comboBox1.Items.Add(dataABC.Columns[i].ColumnName);
                    comboBox2.Items.Add(dataABC.Columns[i].ColumnName);
                    comboBox3.Items.Add(dataABC.Columns[i].ColumnName);
                }
            }
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton5.Checked == false)
            {
                try
                {
                    String name = "Кредитный портфель";
                    String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    "1.xlsx" +
                                    ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

                    OleDbConnection con = new OleDbConnection(constr);
                    OleDbCommand oconn = new OleDbCommand("Select * From [" + name + "$]", con);
                    con.Open();

                    OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                    DataTable dataABC = new DataTable();
                    dataABC.Clear();
                    sda.Fill(dataABC);

                    comboBox1.Items.Clear();
                    comboBox2.Items.Clear();
                    comboBox3.Items.Clear();
                    comboBox1.Text = dataABC.Columns[63].ColumnName;
                    comboBox2.Text = dataABC.Columns[15].ColumnName;
                    comboBox3.Text = dataABC.Columns[21].ColumnName;
                    for (int i = 0; i < dataABC.Columns.Count; i++)
                    {
                        comboBox1.Items.Add(dataABC.Columns[i].ColumnName);
                        comboBox2.Items.Add(dataABC.Columns[i].ColumnName);
                        comboBox3.Items.Add(dataABC.Columns[i].ColumnName);
                    }
                }
                catch { }
            }

            else
            {
                try
                {
                    String name = "licschkre";
                    String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    "2.xlsx" +
                                    ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

                    OleDbConnection con = new OleDbConnection(constr);
                    OleDbCommand oconn = new OleDbCommand("Select * From [" + name + "$]", con);
                    con.Open();

                    OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                    DataTable dataABC = new DataTable();
                    dataABC.Clear();
                    sda.Fill(dataABC);

                    comboBox1.Items.Clear();
                    comboBox2.Items.Clear();
                    comboBox3.Items.Clear();
                    comboBox1.Text = dataABC.Columns[1].ColumnName;
                    comboBox2.Text = dataABC.Columns[39].ColumnName;
                    comboBox3.Text = dataABC.Columns[53].ColumnName;
                    for (int i = 0; i < dataABC.Columns.Count; i++)
                    {
                        comboBox1.Items.Add(dataABC.Columns[i].ColumnName);
                        comboBox2.Items.Add(dataABC.Columns[i].ColumnName);
                        comboBox3.Items.Add(dataABC.Columns[i].ColumnName);
                    }
                }
                catch { }
            }
        }

        private void comboBox4_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try { comboBox4.Text = comboBox4.Text.ToUpper(MyChange.DilDeyisme); }
                catch { }

                try
                {
                    myrefresh();
                }
                catch { }
            }
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                myrefresh();
            }
            catch { }
        }

        private void excelToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            String name = "licschkre";
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            "2.xlsx" +
                            ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(constr);
            OleDbCommand oconn = new OleDbCommand("Select [Ö#G#], [Layihe], Adı, [Lizinqin məbləği], [Qalıq], [% məbləği], [V#K#Qalıq], [V#K#% məbləği], [Cərimə % məbləği], [Dəbbə məbləği], [Son əməl#tarixi], [Qrafikda olan məbləğ], [overdueday] From [" + name + "$] WHERE Not Qalıq=0 or Not [V#K#Qalıq]=0 or Not [V#K#% məbləği]=0 or Not [Dəbbə məbləği]= 0 or Not [Cərimə % məbləği]=0 or Not [% məbləği]=0", con);
            con.Open();

            OleDbDataAdapter abc = new OleDbDataAdapter(oconn);
            DataTable dataABC = new DataTable();
            dataABC.Clear();
            abc.Fill(dataABC);
            ///////////////////datatableni sortlamaq ucun
            dataABC.DefaultView.Sort = "Layihe desc";
            dataABC = dataABC.DefaultView.ToTable(true);
            //dataGridView1.DataSource = data;
            con.Close();

            try
            {
                int s, k, a = 0;

                try { File.Copy("Portfel.xlsm", "C:\\Users\\" + Environment.UserName + "\\Desktop\\PORTFEL (L+S).xlsm", true); }
                catch { MessageBox.Show("'Portfel.xlsm' tapılmadı."); }

                //Get a new workbook.
                oXL = new Excel.Application();
                oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\PORTFEL (L+S).xlsm"));
                oSheet = (Excel._Worksheet)oWB.Sheets[1];

                oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                oXL.Visible = true;
                oSheet.Activate();
                oSheet.Range["A1"].Select();
                oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;

                //oSheet.Cells.EntireColumn.NumberFormat = "@";

                for (s = 0; s < dataABC.Columns.Count; s++)
                {

                    a = s + 2;

                    oSheet.Cells[1, a] = dataABC.Columns[s].ColumnName;
                    oSheet.Cells[1, 1] = "Nö";

                    for (k = 0; k < dataABC.Rows.Count; k++)
                    {
                        oSheet.Cells[k + 2, a] = dataABC.Rows[k][s].ToString();
                        if (dataABC.Columns[s].ColumnName == "Son əməl#tarixi") oSheet.Cells[k + 2, a] = "'" + dataABC.Rows[k][s].ToString();
                        //oSheet.Cells[k + 1, s + 1].Borders.LineStyle = Excel.Constants.xlSolid;

                        if (k != dataABC.Rows.Count) oSheet.Cells[k + 2, 1] = k + 1;
                    }


                }

                DateTime today = DateTime.Now;
                oSheet.Cells[1, 4] = "Adı                                                                                       " + today.ToShortDateString();

                double cem1 = 0, cem2 = 0, cem3 = 0, cem4 = 0, cem5 = 0, cem6 = 0;
                for (k = 0; k < dataABC.Rows.Count; k++)
                {
                    cem1 += Convert.ToDouble(dataABC.Rows[k]["Qalıq"]);
                    cem2 += Convert.ToDouble(dataABC.Rows[k]["% məbləği"]);
                    cem3 += Convert.ToDouble(dataABC.Rows[k]["V#K#Qalıq"]);
                    cem4 += Convert.ToDouble(dataABC.Rows[k]["V#K#% məbləği"]);
                    cem5 += Convert.ToDouble(dataABC.Rows[k]["Cərimə % məbləği"]);
                    cem6 += Convert.ToDouble(dataABC.Rows[k]["Dəbbə məbləği"]);
                }

                oSheet.Cells[dataABC.Rows.Count + 2, 6] = cem1.ToString();
                oSheet.Cells[dataABC.Rows.Count + 2, 7] = cem2.ToString();
                oSheet.Cells[dataABC.Rows.Count + 2, 8] = cem3.ToString();
                oSheet.Cells[dataABC.Rows.Count + 2, 9] = cem4.ToString();
                oSheet.Cells[dataABC.Rows.Count + 2, 10] = cem5.ToString();
                oSheet.Cells[dataABC.Rows.Count + 2, 11] = cem6.ToString();
                oSheet.Cells[dataABC.Rows.Count + 2, 12] = (cem1 + cem2 + cem3 + cem4 + cem5 + cem6).ToString();
                oSheet.Range[oSheet.Cells[dataABC.Rows.Count + 2, 12], oSheet.Cells[dataABC.Rows.Count + 2, 14]].Merge();

                oSheet.Cells[dataABC.Rows.Count + 2, 6].Font.Bold = true;
                oSheet.Cells[dataABC.Rows.Count + 2, 7].Font.Bold = true;
                oSheet.Cells[dataABC.Rows.Count + 2, 8].Font.Bold = true;
                oSheet.Cells[dataABC.Rows.Count + 2, 9].Font.Bold = true;
                oSheet.Cells[dataABC.Rows.Count + 2, 10].Font.Bold = true;
                oSheet.Cells[dataABC.Rows.Count + 2, 11].Font.Bold = true;
                oSheet.Cells[dataABC.Rows.Count + 2, 12].Font.Bold = true;
                oSheet.Columns.AutoFit();
                oSheet.Rows.AutoFit();

                oXL.DisplayAlerts = false;
                oWB.Save();
                //  oSheet.PrintOut();
                //  oWB.Close(SaveChanges: false);
                //  oXL.Application.Quit();        
            }
            catch (System.Exception excep)
            {
                MessageBox.Show(excep.Message + " Məlumatlar Excel-ə doldurularkən, Excel-ə toxunmaq olmaz !!! ");
            }
        }

        public void wordToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string SiyahiPortfel = "";
            DateTime dt = DateTime.Now;

            try { File.Copy("New Emphty.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Gecikme - " + dt.ToShortDateString() + ".doc", true); }
            catch { MessageBox.Show("'\\192.168.10.5\\Common\\AGLizinq\\New Emphty.doc' tapılmadı."); }

            object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\Gecikme - " + dt.ToShortDateString() + ".doc";

            string tarix = "", OdenisGunu = "";

            String name = "licschkre";
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            "2.xlsx" +
                            ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(constr);
            OleDbCommand oconn = new OleDbCommand("Select Adı, [V#K#Qalıq], [V#K#% məbləği], [Dəbbə məbləği], [Cərimə % məbləği], [Qrafikda olan məbləğ], [Verilmə tarixi], [overdueday], [Qalıq], [% məbləği], Layihe, [K#p#b#tarixi], [Son əməl#tarixi], [Uzadılma tarixi] From [" + name + "$] WHERE Not Qalıq=0 or Not [V#K#Qalıq]=0 or Not [V#K#% məbləği]=0 or Not [Dəbbə məbləği]= 0 or Not [Cərimə % məbləği]=0 or Not [% məbləği]=0", con);
            con.Open();
            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
            DataTable data1 = new DataTable();
            data1.Clear();
            sda.Fill(data1);
            con.Close();

            int qenaetbexsSay = 0, nezaretAltindaSay = 0, umidsizSay = 0;
            double CemiGecikme = 0, Gecikme100 = 0, CemiQaliq = 0, NisbetFaiz = 0, k = 0, k2 = 0, qenaetbexs = 0, nezaretAltinda = 0, umidsiz = 0;
            string SonOdeme = "20150101", Vaxt1 = "", Vaxt2 = "", Vaxt3 = "", Vaxt4 = ""; 

            for (int i = 0; i < data1.Rows.Count; i++)
            {
                try
                {
                    SonOdeme = data1.Rows[i]["Son əməl#tarixi"].ToString().Substring(6, 4) + data1.Rows[i]["Son əməl#tarixi"].ToString().Substring(3, 2) + data1.Rows[i]["Son əməl#tarixi"].ToString().Substring(0, 2);
                }
                catch { SonOdeme = "20150101"; }
                
                Vaxt1 = dt.AddDays(-90).ToShortDateString();
                Vaxt2 = Vaxt1.Substring(6, 4) + Vaxt1.Substring(3, 2) + Vaxt1.Substring(0, 2);

                Vaxt3 = dt.AddDays(-30).ToShortDateString();
                Vaxt4 = Vaxt3.Substring(6, 4) + Vaxt3.Substring(3, 2) + Vaxt3.Substring(0, 2);

                if (Convert.ToInt32(SonOdeme) > Convert.ToInt32(Vaxt4))
                {
                    qenaetbexsSay += 1;
                    qenaetbexs += Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"]) + Convert.ToDouble(data1.Rows[i]["Qalıq"]) + Convert.ToDouble(data1.Rows[i]["% məbləği"]);
                }

                if (Convert.ToInt32(SonOdeme) > Convert.ToInt32(Vaxt2) && Convert.ToInt32(SonOdeme) <= Convert.ToInt32(Vaxt4))
                {
                    nezaretAltindaSay += 1;
                    nezaretAltinda += Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"]) + Convert.ToDouble(data1.Rows[i]["Qalıq"]) + Convert.ToDouble(data1.Rows[i]["% məbləği"]);
                }

                if (Convert.ToInt32(SonOdeme) <= Convert.ToInt32(Vaxt2))
                {
                    umidsizSay += 1;
                    umidsiz += Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"]) + Convert.ToDouble(data1.Rows[i]["Qalıq"]) + Convert.ToDouble(data1.Rows[i]["% məbləği"]);
                }

                if (data1.Rows[i]["Uzadılma tarixi"].ToString() != "") OdenisGunu = data1.Rows[i]["Uzadılma tarixi"].ToString().Substring(0, 2);
                else OdenisGunu = data1.Rows[i]["K#p#b#tarixi"].ToString().Substring(0, 2);

                CemiQaliq += Math.Round(Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"]) + Convert.ToDouble(data1.Rows[i]["Qalıq"]) + Convert.ToDouble(data1.Rows[i]["% məbləği"]), 2);


                //// BEZI MUSTERILER UCUN XUSUSI YAZILIM ###### GECIKMIS FAIZ BORCUN GORUNMEMESI UCUN #########.........................
                if (data1.Rows[i]["Layihe"].ToString() == "'S-008/14" || data1.Rows[i]["Layihe"].ToString() == "'S-01/20-058/15" || data1.Rows[i]["Layihe"].ToString() == "'S-058/15") { SiyahiPortfel += Environment.NewLine + OdenisGunu + " ♦ " + data1.Rows[i]["Adı"].ToString() + " - " + Math.Round((Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"])), 2).ToString() + "+" + Math.Round(Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]), 2) + "/" + Math.Round((Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"])), 2).ToString() + " ♦ aylıq - " + data1.Rows[i]["Qrafikda olan məbləğ"].ToString() + " ♦ " + data1.Rows[i]["Layihe"].ToString().Substring(1, data1.Rows[i]["Layihe"].ToString().Length - 1) + " ♦ son ödəmə - " + data1.Rows[i]["Son əməl#tarixi"].ToString(); }
                else { SiyahiPortfel += Environment.NewLine + OdenisGunu + " ♦ " + data1.Rows[i]["Adı"].ToString() + " - " + Math.Round((Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"])), 2).ToString() + "/" + Math.Round((Convert.ToDouble(data1.Rows[i]["% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"])), 2).ToString() + " ♦ aylıq - " + data1.Rows[i]["Qrafikda olan məbləğ"].ToString() + " ♦ " + data1.Rows[i]["Layihe"].ToString().Substring(1, data1.Rows[i]["Layihe"].ToString().Length - 1) + " ♦ son ödəmə - " + data1.Rows[i]["Son əməl#tarixi"].ToString(); }
                
                if ((Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"]) > 0))
                {
                    k += 1;
                    CemiGecikme += Math.Round(Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"]), 2);
                }

                if (Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"]) > 0 && Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"]) < 100)
                {
                    k2 += 1;
                    Gecikme100 += Math.Round(Convert.ToDouble(data1.Rows[i]["V#K#Qalıq"]) + Convert.ToDouble(data1.Rows[i]["V#K#% məbləği"]) + Convert.ToDouble(data1.Rows[i]["Dəbbə məbləği"]) + Convert.ToDouble(data1.Rows[i]["Cərimə % məbləği"]), 2);
                }
            }
            //notifyIcon1.Icon = SystemIcons.;
            NisbetFaiz =Convert.ToInt32(CemiGecikme / CemiQaliq * 100); //gecikmenin faizinin tapilmasi
            //pertfel vaxti ucun
            try
            {
                String name2 = "licschkre";
                String constr2 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                "2.xlsx" +
                                ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

                OleDbConnection con2 = new OleDbConnection(constr2);
                OleDbCommand oconn2 = new OleDbCommand("Select overdueday From [" + name2 + "$]", con2);
                con2.Open();
                OleDbDataAdapter sda2 = new OleDbDataAdapter(oconn2);
                DataTable data2 = new DataTable();
                data2.Clear();
                sda2.Fill(data2);
                con2.Close();
                tarix = "PORTFEL TARİXİ ♦ " + data2.Rows[0]["overdueday"].ToString();
            }
            catch { }

            string qeydler = "";

            MyData.selectCommand("baza.accdb", "Select * from Qeydler");
            MyData.dtmain= new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmain);

            if (MyData.dtmain.Rows.Count != 0)
            {
                for (int i = 0; i < MyData.dtmain.Rows.Count; i++)
                {
                    qeydler += Environment.NewLine + MyData.dtmain.Rows[i][1].ToString().Substring(0, 10) + " - " + MyData.dtmain.Rows[i][2].ToString();
                }

                qeydler = "QEYDLƏR:" + qeydler;
            }

            Microsoft.Office.Interop.Word._Application oWord;
            object oMissing = Type.Missing;
            oWord = new Microsoft.Office.Interop.Word.Application();
            oWord.Visible = true;
            oWord.Documents.Open(FileName);
            oWord.Selection.TypeText(tarix + Environment.NewLine + SiyahiPortfel + Environment.NewLine + Environment.NewLine + "Ümumi (100%) - " + data1.Rows.Count.ToString() + " ədəd (" + CemiQaliq + " AZN)" + Environment.NewLine + "Gecikmiş (" + NisbetFaiz.ToString() + "%) - " + k.ToString() + " ədəd (" + CemiGecikme + " AZN)" + Environment.NewLine + Environment.NewLine + "Qənaətbəxş (" + Convert.ToInt32(qenaetbexs * 100 / (qenaetbexs + nezaretAltinda + umidsiz)) + "%) - " + qenaetbexsSay.ToString() + " ədəd (" + qenaetbexs.ToString() + " AZN)" + Environment.NewLine + "Nəzarət altında olan (" + Convert.ToInt32(nezaretAltinda * 100 / (qenaetbexs + nezaretAltinda + umidsiz)) + "%) - " + nezaretAltindaSay.ToString() + " ədəd (" + nezaretAltinda.ToString() + " AZN)" + Environment.NewLine + "Ümidsiz (" + Convert.ToInt32(umidsiz * 100 / (qenaetbexs + nezaretAltinda + umidsiz)) + "%) - " + umidsizSay.ToString() + " ədəd (" + umidsiz.ToString() + " AZN)"); 
            oWord.ActiveDocument.Save();
            //oWord.Quit();
            //MessageBox.Show("The text is inserted.");

            //}
            //catch { MessageBox.Show("Səhv var."); }
        }

        private void excelÜmumiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            String name = "licschkre";
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            "2.xlsx" +
                            ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(constr);
            OleDbCommand oconn = new OleDbCommand("Select [Ö#G#], [Layihe], Adı, [Lizinqin məbləği], [Qalıq], [% məbləği], [V#K#Qalıq], [V#K#% məbləği], [Cərimə % məbləği], [Dəbbə məbləği], [Son əməl#tarixi], [Qrafikda olan məbləğ], [overdueday] From [" + name + "$]", con);
            con.Open();

            OleDbDataAdapter abc = new OleDbDataAdapter(oconn);
            DataTable dataABC = new DataTable();
            dataABC.Clear();
            abc.Fill(dataABC);
            ///////////////////datatableni sortlamaq ucun
            dataABC.DefaultView.Sort = "Layihe desc";
            dataABC = dataABC.DefaultView.ToTable(true);
            //dataGridView1.DataSource = data;
            con.Close();

            try
            {
                int s, k, a = 0;

                try { File.Copy("Portfel.xlsm", "C:\\Users\\" + Environment.UserName + "\\Desktop\\PORTFEL (L+S).xlsm", true); }
                catch { MessageBox.Show("'Portfel.xlsm' tapılmadı."); }

                //Get a new workbook.
                oXL = new Excel.Application();
                oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\PORTFEL (L+S).xlsm"));
                oSheet = (Excel._Worksheet)oWB.Sheets[1];

                oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                oXL.Visible = true;
                oSheet.Activate();
                oSheet.Range["A1"].Select();
                oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;

                //oSheet.Cells.EntireColumn.NumberFormat = "@";

                for (s = 0; s < dataABC.Columns.Count; s++)
                {

                    a = s + 2;

                    oSheet.Cells[1, a] = dataABC.Columns[s].ColumnName;
                    oSheet.Cells[1, 1] = "Nö";

                    for (k = 0; k < dataABC.Rows.Count; k++)
                    {
                        oSheet.Cells[k + 2, a] = dataABC.Rows[k][s].ToString();
                        if (dataABC.Columns[s].ColumnName == "Son əməl#tarixi") oSheet.Cells[k + 2, a] = "'" + dataABC.Rows[k][s].ToString();
                        //oSheet.Cells[k + 1, s + 1].Borders.LineStyle = Excel.Constants.xlSolid;

                        if (k != dataABC.Rows.Count) oSheet.Cells[k + 2, 1] = k + 1;
                    }
                    
                }

                DateTime today = DateTime.Now;
                oSheet.Cells[1, 4] = "Adı                                                                                       " + today.ToShortDateString();

                double cem1 = 0, cem2 = 0, cem3 = 0, cem4 = 0, cem5 = 0, cem6 = 0;
                for (k = 0; k < dataABC.Rows.Count; k++)
                {
                    cem1 += Convert.ToDouble(dataABC.Rows[k]["Qalıq"]);
                    cem2 += Convert.ToDouble(dataABC.Rows[k]["% məbləği"]);
                    cem3 += Convert.ToDouble(dataABC.Rows[k]["V#K#Qalıq"]);
                    cem4 += Convert.ToDouble(dataABC.Rows[k]["V#K#% məbləği"]);
                    cem5 += Convert.ToDouble(dataABC.Rows[k]["Cərimə % məbləği"]);
                    cem6 += Convert.ToDouble(dataABC.Rows[k]["Dəbbə məbləği"]);
                }

                oSheet.Cells[dataABC.Rows.Count + 2, 6] = cem1.ToString();
                oSheet.Cells[dataABC.Rows.Count + 2, 7] = cem2.ToString();
                oSheet.Cells[dataABC.Rows.Count + 2, 8] = cem3.ToString();
                oSheet.Cells[dataABC.Rows.Count + 2, 9] = cem4.ToString();
                oSheet.Cells[dataABC.Rows.Count + 2, 10] = cem5.ToString();
                oSheet.Cells[dataABC.Rows.Count + 2, 11] = cem6.ToString();
                oSheet.Cells[dataABC.Rows.Count + 2, 12] = (cem1 + cem2 + cem3 + cem4 + cem5 + cem6).ToString();
                oSheet.Range[oSheet.Cells[dataABC.Rows.Count + 2, 12], oSheet.Cells[dataABC.Rows.Count + 2, 14]].Merge();

                oSheet.Cells[dataABC.Rows.Count + 2, 6].Font.Bold = true;
                oSheet.Cells[dataABC.Rows.Count + 2, 7].Font.Bold = true;
                oSheet.Cells[dataABC.Rows.Count + 2, 8].Font.Bold = true;
                oSheet.Cells[dataABC.Rows.Count + 2, 9].Font.Bold = true;
                oSheet.Cells[dataABC.Rows.Count + 2, 10].Font.Bold = true;
                oSheet.Cells[dataABC.Rows.Count + 2, 11].Font.Bold = true;
                oSheet.Cells[dataABC.Rows.Count + 2, 12].Font.Bold = true;
                oSheet.Columns.AutoFit();
                oSheet.Rows.AutoFit();

                oXL.DisplayAlerts = false;
                oWB.Save();
                //  oSheet.PrintOut();
                //  oWB.Close(SaveChanges: false);
                //  oXL.Application.Quit();        
            }
            catch (System.Exception excep)
            {
                MessageBox.Show(excep.Message + " Məlumatlar Excel-ə doldurularkən, Excel-ə toxunmaq olmaz !!! ");
            }
        }

    }
}
