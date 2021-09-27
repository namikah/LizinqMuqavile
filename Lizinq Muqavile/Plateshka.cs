using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Windows.Forms;
using Nsoft;

namespace Lizinq_Muqavile
{
    public partial class Plateshka : Form
    {
        public Plateshka()
        {
            InitializeComponent();
        }

        Excel.Application oXL;
        Excel._Workbook oWB;
        Excel._Worksheet oSheet;

        private void anonsrefresh()
        {
            DateTime dt = dttarix.Value.Date.AddYears(-1);
            string s = "";
            lbAnons.Text = "Keçən il bu ay: ♦ ";

            MyData.selectCommand("baza.accdb", "SELECT * FROM plateshkaARXIV WHERE TARİX like '%" + dt.Month + "-" + dt.Year + "%'");
            MyData.dtmainArxiv=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainArxiv);

            for (int i = 0; i < MyData.dtmainArxiv.Rows.Count; i++)
            {
                s += MyData.dtmainArxiv.Rows[i]["HARA"].ToString() + " - " + MyData.dtmainArxiv.Rows[i]["MƏBLƏĞ"].ToString() + " AZN ♦ ";

            }

            lbAnons.Text = s;
        }

        private void emekhaqqirefresh()  //---------------rasxodun nomresinin load olunmasi-------------------------------
        {
            try
            {
                double e1 = 0, e2 = 0, e3 = 0, e4 = 0, e5 = 0, e6 = 0, e7 = 0;
                
                MyData.selectCommand("baza.accdb", "Select * From Emekhaqqi");
                MyData.dtmainemekhaqqi=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainemekhaqqi);
                dataGridView2.DataSource = MyData.dtmainemekhaqqi;

                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {
                    try { e1 += Math.Round(Convert.ToDouble(dataGridView2.Rows[i].Cells["c4"].Value.ToString()), 2, MidpointRounding.AwayFromZero); }
                    catch { }
                    try { e2 += Math.Round(Convert.ToDouble(dataGridView2.Rows[i].Cells["c5"].Value.ToString()), 2, MidpointRounding.AwayFromZero); }
                    catch { }
                    try { e3 += Math.Round(Convert.ToDouble(dataGridView2.Rows[i].Cells["c6"].Value.ToString()), 2, MidpointRounding.AwayFromZero); }
                    catch { }
                    try { e4 += Math.Round(Convert.ToDouble(dataGridView2.Rows[i].Cells["c7"].Value.ToString()), 2, MidpointRounding.AwayFromZero); }
                    catch { }
                    try { e5 += Math.Round(Convert.ToDouble(dataGridView2.Rows[i].Cells["c8"].Value.ToString()), 2, MidpointRounding.AwayFromZero); }
                    catch { }
                    try { e6 += Math.Round(Convert.ToDouble(dataGridView2.Rows[i].Cells["c88"].Value.ToString()), 2, MidpointRounding.AwayFromZero); }
                    catch { }
                    try { e7 += Math.Round(Convert.ToDouble(dataGridView2.Rows[i].Cells["c4"].Value.ToString()) / 100, 2, MidpointRounding.AwayFromZero); }
                    catch { }
                }

                btEmekHaqqi2.Text = e1.ToString();
                btGelirVergisi2.Text = e2.ToString();
                btPensiya3Faiz2.Text = e3.ToString();
                btPensiya22Faiz2.Text = e4.ToString();
                btIwsizlik05Faiz2.Text = Math.Round(Convert.ToDouble(e5) * 2, 2, MidpointRounding.AwayFromZero).ToString();
                btTibbiSigorta2.Text = Math.Round(Convert.ToDouble(e6) * 2, 2, MidpointRounding.AwayFromZero).ToString();
                btNagdilasdirma2.Text = Math.Round(Convert.ToDouble(e7), 2, MidpointRounding.AwayFromZero).ToString();

                txtmanatEmekhaqqi.Text = e1.ToString();
                txtmanat3faiz.Text = e3.ToString();
                txtmanat22faiz.Text = e4.ToString();
                txtmanatgelirvergisi.Text = e2.ToString();
                txtmanatSigEden05faiz.Text = Math.Round(Convert.ToDouble(e5), 2, MidpointRounding.AwayFromZero).ToString();
                txtmanatSigolunan05faiz.Text = Math.Round(Convert.ToDouble(e5), 2, MidpointRounding.AwayFromZero).ToString();
                txtmanatIsegoturen1faiz.Text = Math.Round(Convert.ToDouble(e6), 2, MidpointRounding.AwayFromZero).ToString();
                txtmanatIsciler1faiz.Text = Math.Round(Convert.ToDouble(e6), 2, MidpointRounding.AwayFromZero).ToString();

                listBox3.Items.Clear();
                listBox3.Items.Add("Ə. haqqi -- " + e1.ToString());
                listBox3.Items.Add("DSMF  10% -- " + e3.ToString());
                listBox3.Items.Add("DSMF 15% -- " + e4.ToString());
                listBox3.Items.Add("İTS 2% -- " + Math.Round(Convert.ToDouble(e6) * 2, 2, MidpointRounding.AwayFromZero).ToString());
                listBox3.Items.Add("İSH  0.5% -- " + Math.Round(Convert.ToDouble(e5) * 2, 2, MidpointRounding.AwayFromZero).ToString());
            }
            catch { }
        }

        private void ARXIV()  //---------------rasxodun nomresinin load olunmasi------------------------------------------
        {
            string commandText = "SELECT * FROM plateshkaARXIV WHERE 1=1";
            commandText += " and TARİX between #" + dtBaslama.Value.ToString("yyyy-MM-dd") + "# and #" + dtBitme.Value.AddDays(1).ToString("yyyy-MM-dd") + "#";
            commandText += " and HARA like '%" + txtAxtarHara.Text + "%'";
            commandText += " and MƏBLƏĞ like '%" + txtAxtarMebleg.Text + "%'";
            commandText += " and [ÖDƏNİŞİN TƏYİNATI] like '%" + txtAxtarTeyinat.Text + "%'";
            commandText += " order by Код desc";

            MyData.selectCommand("baza.accdb", commandText);
            MyData.dtmainArxiv = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainArxiv);
            dataGridView1.DataSource = MyData.dtmainArxiv;

            //this.dataGridView1.Sort(this.dataGridView1.Columns["Код"], ListSortDirection.Descending);

            try
            {
                double k = 0;
                for (int i = 0; i < dataGridView1.Rows.Count; i++) k = Math.Round(Convert.ToDouble(k) + Convert.ToDouble(dataGridView1.Rows[i].Cells["MƏBLƏĞ"].Value), 2, MidpointRounding.AwayFromZero);
                btcemi.Text = "Cəmi: " + k.ToString();
            }
            catch { }
        }

        private void CemRefresh()
        {
            btCem.Text = Math.Round((Convert.ToDouble(btEmekHaqqi2.Text) + (Convert.ToDouble(btNagdilasdirma2.Text) + Convert.ToDouble(btGelirVergisi2.Text) + Convert.ToDouble(btPensiya3Faiz2.Text) + Convert.ToDouble(btPensiya22Faiz2.Text) + Convert.ToDouble(btIwsizlik05Faiz2.Text) + Convert.ToDouble(btTibbiSigorta2.Text))),2,MidpointRounding.AwayFromZero).ToString();
        }

        private void IndividualHesablamaRefresh()
        {
                MyData.updateCommand("baza.accdb", "UPDATE Emekhaqqi SET "
                                                                                     + "c1 ='" + dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["c1"].Value.ToString() + "',"
                                                                                     + "c11 ='" + dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["c11"].Value.ToString() + "',"
                                                                                     + "c2 ='" + dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["c2"].Value.ToString() + "',"
                                                                                     + "c3 ='" + dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["c3"].Value.ToString() + "',"
                                                                                     + "c9 ='" + dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["c9"].Value.ToString() + "',"
                                                                                     + "c5 ='" + dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["c5"].Value.ToString() + "',"
                                                                                     + "c6 ='" + dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["c6"].Value.ToString() + "',"
                                                                                     + "c7 ='" + dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["c7"].Value.ToString() + "',"
                                                                                     + "c8 ='" + dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["c8"].Value.ToString() + "',"
                                                                                     + "c88 ='" + dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["c88"].Value.ToString() + "',"
                                                                                     + "c10 ='" + dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["c10"].Value.ToString() + "',"
                                                                                     + "c4 ='" + dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["c4"].Value.ToString() + "'"
                                                                                     + " WHERE Kod Like '" + dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["Kod"].Value.ToString() + "'");

                
                
        }

        private void HesablamaRefresh()
        {
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                double CemiEmekhaqqi = Convert.ToDouble(dataGridView2.Rows[i].Cells["c2"].Value) + Convert.ToDouble(dataGridView2.Rows[i].Cells["c3"].Value) + Convert.ToDouble(dataGridView2.Rows[i].Cells["c9"].Value);
                double GelirVergisi = 0; if (checkBox11.Checked == true) GelirVergisi = Math.Round((CemiEmekhaqqi - 200) * 14 / 100, 2, MidpointRounding.AwayFromZero);
                double İcbariSiğorta = 0; if (checkBox10.Checked == true) İcbariSiğorta = Math.Round((CemiEmekhaqqi * 2 / 100) * 50 / 100, 2, MidpointRounding.AwayFromZero); // 2%-ə 50% Guzest
                double Faiz10 = Math.Round(6 + (CemiEmekhaqqi - 200) * 10 / 100, 2, MidpointRounding.AwayFromZero); // sosial sigorta
                double Faiz15 = Math.Round(44 + (CemiEmekhaqqi - 200) * 15 / 100, 2, MidpointRounding.AwayFromZero);  // sosial sigorta
                double Faiz05 = Math.Round(CemiEmekhaqqi * 5 / 1000, 2, MidpointRounding.AwayFromZero); //iwsizlikden sigorta
                double CemiTutulma = Faiz10 + Faiz05 + İcbariSiğorta + GelirVergisi;
                double NetEmekhaqqi = Convert.ToDouble(CemiEmekhaqqi - CemiTutulma);
                //double BirFaiz = Convert.ToDouble(NetEmekhaqqi / 100); // LAZIM OLMADI DEYE ISTIFADE OLUNMUR....

                
                MyData.updateCommand("baza.accdb", "UPDATE Emekhaqqi SET "
                                                                                     + "c1 ='" + dataGridView2.Rows[i].Cells["c1"].Value.ToString() + "',"
                                                                                     + "c11 ='" + dataGridView2.Rows[i].Cells["c11"].Value.ToString() + "',"
                                                                                     + "c2 ='" + dataGridView2.Rows[i].Cells["c2"].Value.ToString() + "',"
                                                                                     + "c3 ='" + dataGridView2.Rows[i].Cells["c3"].Value.ToString() + "',"
                                                                                     + "c9 ='" + dataGridView2.Rows[i].Cells["c9"].Value.ToString() + "',"
                                                                                     + "c5 ='" + Math.Round(GelirVergisi, 2, MidpointRounding.AwayFromZero).ToString() + "',"
                                                                                     + "c6 ='" + Math.Round(Faiz10, 2, MidpointRounding.AwayFromZero).ToString() + "'," //evveller 10% evezine 3 faiz
                                                                                     + "c7 ='" + Math.Round(Faiz15, 2, MidpointRounding.AwayFromZero).ToString() + "'," //evveller 15% evezine 22 faiz
                                                                                     + "c8 ='" + Math.Round(Faiz05, 2, MidpointRounding.AwayFromZero).ToString() + "',"
                                                                                     + "c88 ='" + Math.Round(İcbariSiğorta, 2, MidpointRounding.AwayFromZero).ToString() + "',"
                                                                                     + "c10 ='" + Math.Round(CemiTutulma, 2, MidpointRounding.AwayFromZero).ToString() + "',"
                                                                                     + "c4 = '" + Math.Round((NetEmekhaqqi), 2, MidpointRounding.AwayFromZero) + "'"
                                                                                     + " WHERE Kod Like '" + dataGridView2.Rows[i].Cells["Kod"].Value.ToString() + "'");

                
                
            }
        }

        private void reqemler()      //------reqem yazi ile---------------------------------------------------------------
        {
            try
            {
                txt2.Text = MyChange.ReqemToMetn(Convert.ToDouble(txtmanat.Text));
            }
            catch { }
        }

        private void reqemler2()      //------reqem yazi ile--------------------------------------------------------------
        {
            try
            {
                txtedvherfle.Text = MyChange.ReqemToMetn(Convert.ToDouble(txtedvmanat.Text));
            }
            catch { }
        }

        private void reqemler3()       //------reqem yazi ile-------------------------------------------------------------
        {
            try
            {
                txtherf5.Text = MyChange.ReqemToMetn(Convert.ToDouble(txtkonvertmanat.Text));
            }
            catch { }
        }

        private void reqemler4()      //------reqem yazi ile--------------------------------------------------------------
        {
            try
            {
                txtmanatEmekhaqqi2.Text = MyChange.ReqemToMetn(Convert.ToDouble(txtmanatEmekhaqqi.Text));
            }
            catch { }
        }

        private void reqemler5()      //------reqem yazi ile--------------------------------------------------------------
        {
            try
            {
                txtmanatEmekhaqqi4.Text = MyChange.ReqemToMetn(Convert.ToDouble(txtmanat3faiz.Text));
            }
            catch { }
        }

        private void reqemler6()      //------reqem yazi ile--------------------------------------------------------------
        {
            try
            {
                txtmanatEmekhaqqi6.Text = MyChange.ReqemToMetn(Convert.ToDouble(txtmanat22faiz.Text));
            }
            catch { }
        }

        private void reqemler7()      //------reqem yazi ile--------------------------------------------------------------
        {
            try
            {
                txtmanatEmekhaqqi8.Text = MyChange.ReqemToMetn(Convert.ToDouble(txtmanatgelirvergisi.Text));
            }
            catch { }
        }

        private void reqemler8()      //------reqem yazi ile--------------------------------------------------------------
        {
            try
            {
                txtmanatEmekhaqqi10.Text = MyChange.ReqemToMetn(Convert.ToDouble(txtmanatSigEden05faiz.Text));
            }
            catch { }
        }

        private void reqemler9()      //------reqem yazi ile--------------------------------------------------------------
        {
            try
            {
                txtmanatEmekhaqqi12.Text = MyChange.ReqemToMetn(Convert.ToDouble(txtmanatSigolunan05faiz.Text));
            }
            catch { }
        }

        private void reqemler10()      //------reqem yazi ile--------------------------------------------------------------
        {
            try
            {
                txtmanatIsegoturen1faiz2.Text = MyChange.ReqemToMetn(Convert.ToDouble(txtmanatIsegoturen1faiz.Text));
            }
            catch { }
        }

        private void reqemler11()      //------reqem yazi ile--------------------------------------------------------------
        {
            try
            {
                txtmanatIsciler1faiz2.Text = MyChange.ReqemToMetn(Convert.ToDouble(txtmanatIsciler1faiz.Text));
            }
            catch { }
        }

        private void rekvizitlerRefresh()  //---------------rasxodun nomresinin load olunmasi--------------------------------------
        {
            comboBox1.Items.Clear();
            comboBox2.Items.Clear();
            cbFromEmekHaqqi.Items.Clear();
            int k;
            
            MyData.selectCommand("baza.accdb", "Select * From rekvizitler");
            MyData.dtmainrekvizitler=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

            for (k = 0; k < MyData.dtmainrekvizitler.Rows.Count; k++)
            {
                comboBox1.Items.Add(MyData.dtmainrekvizitler.Rows[k][1].ToString());
                comboBox2.Items.Add(MyData.dtmainrekvizitler.Rows[k][1].ToString());
                cbFromEmekHaqqi.Items.Add(MyData.dtmainrekvizitler.Rows[k][1].ToString());

            }
        }

        private void plateshkanomreRefresh()  //---------------platyoska nomresinin load olunmasi------------------------------------
        {
            MyData.selectCommand("baza.accdb", "Select * From plateshkanomre");
            MyData.dtmainplateshkanomre=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainplateshkanomre);
            try
            {
                txtnomresi.Text = (Convert.ToInt32(MyData.dtmainplateshkanomre.Rows[0][0]) + 1).ToString();
                txtnomreEmekhaqqi.Text = (Convert.ToInt32(MyData.dtmainplateshkanomre.Rows[0][0]) + 1).ToString();
            }
            catch { MessageBox.Show("Məxaric nömrəsində səhv..."); };
        }

        private void EDVRefresh()  //-------------------------------------------------------------------------------------
        {
            cbedv2.Items.Clear();
            int k;

            MyData.selectCommand("baza.accdb", "Select * From plateshkaEDV");
            MyData.dtmainedv=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainedv);

            for (k = 0; k < MyData.dtmainedv.Rows.Count; k++)
            {
                cbedv2.Items.Add(MyData.dtmainedv.Rows[k]["a1"].ToString());

            }

        }

        private void EDVnomreRefresh()  //-------------------------------------------------------------------------------------
        {
            MyData.selectCommand("baza.accdb", "Select * From plateshkaEDVnomre");
            MyData.dtmainpEDVnomre=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainpEDVnomre);
            try
            {
                txtedvnomre.Text = (Convert.ToInt32(MyData.dtmainpEDVnomre.Rows[0][0]) + 1).ToString();
            }
            catch { MessageBox.Show("Məxaric nömrəsində səhv..."); };
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (txtnomresi.Enabled == false || txtnomresi2.Enabled == false) { txtnomresi2.Enabled = true; txtnomresi.Enabled = true; return; }
            else if (txtnomresi.Enabled == true || txtnomresi2.Enabled == true) { txtnomresi.Enabled = false; txtnomresi2.Enabled = false; }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (dttarix.Enabled == false ) { dttarix.Enabled = true; return; }
            else if (dttarix.Enabled == true )  { dttarix.Enabled = false; }
        }

        private void btelaveet_Click(object sender, EventArgs e)
        {
            if (!MyCheck.davamYesNo()) return;

            if (txtmusteriadi.Text == "") { MessageBox.Show("Müştərinin Adı qeyd olunmayıb"); return; }
            if (txthesab1.Text == "") { MessageBox.Show("Müştərinin Hesabı qeyd olunmayıb"); return; }
            if (txthesab2.Text == "") { MessageBox.Show("Müştərinin Hesabı qeyd olunmayıb"); return; }
            if (txtmusterivoen.Text == "") { MessageBox.Show("Müştərinin VÖEN-i qeyd olunmayıb"); return; }
            if (txtbankadi.Text == "") { MessageBox.Show("Bankın Adı qeyd olunmayıb"); return; }
            if (txtbankkodu.Text == "") { MessageBox.Show("Bankın Kodu qeyd olunmayıb"); return; }
            if (txtbankvoen.Text == "") { MessageBox.Show("Bankın VÖEN-i qeyd olunmayıb"); return; }
            if (txtmuxhesab1.Text == "") { MessageBox.Show("Bankın Müxbir Hesabı qeyd olunmayıb"); return; }
            if (txtmuxhesab2.Text == "") { MessageBox.Show("Bankın Müxbir Hesabı qeyd olunmayıb"); return; }
            if (txtsvift.Text == "") { MessageBox.Show("Bankın S.W.I.F.T-i qeyd olunmayıb"); return; } 

            int k2=0;
            
            MyData.selectCommand("baza.accdb", "Select * From rekvizitler");
            MyData.dtmainrekvizitler=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

            for (int k = 0; k < MyData.dtmainrekvizitler.Rows.Count; k++)
            {
                if (txtmusteriadi.Text == MyData.dtmainrekvizitler.Rows[k][1].ToString()) k2 += 1; 
            }

            if (k2 > 0) { MessageBox.Show("Qeyd olunan müştəri artıq mövcuddur."); return; }

            MyData.insertCommand("baza.accdb", "insert into rekvizitler (a1, a2, a3, a4, a5, a6, a7, a8, a9) Values ('" + txtmusteriadi.Text + "'," + "'" + txthesab1.Text + " " + txthesab2.Text + "'," + "'" + txtmusterivoen.Text + "'," + "'" + txtbankadi.Text + "'," + "'" + txtbankkodu.Text + "'," + "'" + txtbankvoen.Text + "'," + "'" + txtmuxhesab1.Text + " " + txtmuxhesab2.Text + "'," + "'" + txtsvift.Text + "'," + "'" + txtsexsiyyet.Text + "')");
            
            MyData.insertCommand("baza.accdb", "insert into plateshkateyinat (a1, a2) Values ('" + txtmusteriadi.Text + "','" + txtesas.Text + "')");
                
            MessageBox.Show("Müştəri əlavə olundu.");

            rekvizitlerRefresh();
        }

        private void Plateshka_Load(object sender, EventArgs e)
        {
            DateTime dt = DateTime.Now;
            string t = MyChange.TarixSozle(dt);

            textBox8.Text = dt.Year.ToString() + " - ci il " + t.ToString() + " ayı üçün gəlir vergisi.";

            txtedvnomre2.Text = dt.Year.ToString().Substring(2, 2);
            txtnomresi2.Text = dt.Year.ToString().Substring(2,2);
            txtnomreEmekhaqqi2.Text = dt.Year.ToString().Substring(2, 2);
            dttarix.Text = dt.ToShortDateString();
            dtedvtarix.Text = dt.ToShortDateString();
            dateTimePicker1.Text = dt.ToShortDateString();
            dtBaslama.Value = dtBitme.Value.AddYears(-1);

            emekhaqqirefresh();
            rekvizitlerRefresh();
            plateshkanomreRefresh();
            EDVRefresh();
            EDVnomreRefresh();
            ARXIV();
            CemRefresh();
            anonsrefresh();

            /*SpeechSynthesizer speech = new SpeechSynthesizer();
            speech.Rate = (int)0;
            speech.Speak("Payment Order");*/
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (!MyCheck.davamYesNo()) return;

            try
            {
                if (txtmanat.Text.Substring(txtmanat.Text.Length - 3, 1) != ".")
                {
                    MessageBox.Show("Məbləğ düzgün qeyd olunmayıb." + Environment.NewLine + "Nümunə: 123.45"); return;
                }
            }
            catch { MessageBox.Show("Məbləğ düzgün qeyd olunmayıb." + Environment.NewLine + "Nümunə: 123.45"); return; }

            button5.Enabled = false;
            button7.Enabled = false;

            try { File.Copy("Plateshka.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Odenis Tapsirigi - " + txtmanat.Text + ".xlsx", true); }
            catch { MessageBox.Show("'Plateshka.xlsx' tapılmadı."); }

            try
            {
                MyData.selectCommand("baza.accdb", "Select * From plateshkaARXIV where №=" + "'" + txtnomresi.Text + "/" + txtnomresi2.Text + "' and HARA='" + comboBox2.Text + "'");
                MyData.dtmainArxiv=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainArxiv);

                if (MyData.dtmainArxiv.Rows.Count > 0)
                {
                    MyData.updateCommand("baza.accdb", "UPDATE plateshkaARXIV SET "
                                                                                         + "TARİX =" + "'" + dttarix.Text + "',"
                                                                                         + "№ =" + "'" + txtnomresi.Text + "/" + txtnomresi2.Text + "',"
                                                                                         + "HARDAN =" + "'" + comboBox1.Text + "',"
                                                                                         + "HARA =" + "'" + comboBox2.Text + "',"
                                                                                         + "[ÖDƏNİŞİN TƏYİNATI] =" + "'" + txtteyinat.Text + "',"
                                                                                         + "MƏBLƏĞ =" + "'" + txtmanat.Text + "',"
                                                                                         + "VALYUTA =" + "'" + comboBox3.Text + "'"
                                                                                         + " WHERE №=" + "'" + txtnomresi.Text + "/" + txtnomresi2.Text + "'");

                    
                    
                    ARXIV();
                }
                else
                {
                    MyData.insertCommand("baza.accdb", "insert into plateshkaARXIV (TARİX, №, HARDAN, HARA, [ÖDƏNİŞİN TƏYİNATI], MƏBLƏĞ, VALYUTA) Values ('" + dttarix.Text + "', '" + txtnomresi.Text + "/" + txtnomresi2.Text + "', '" + comboBox1.Text + "', '" + comboBox2.Text + "', '" + txtteyinat.Text + "', '" + txtmanat.Text + "', '" + comboBox3.Text + "')");
                    
                    ARXIV();
                }

            }
            catch { MessageBox.Show("Arxivə kopyalamadı.."); };

            if (comboBox2.Text == "") { MessageBox.Show("Benefisiar (Alan) müştəri seçilməyib.."); return; }

            DateTime dt = DateTime.Now;
            string a = MyChange.TarixSozle(dt);

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Odenis Tapsirigi - " + txtmanat.Text + ".xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            if (txtmanat.Text == "") txtmanat.Text = "0";

            oSheet.Cells[2, 2] = "ÖDƏNİŞ TAPŞIRIĞI № " + txtnomresi.Text + "/" + txtnomresi2.Text;
            oSheet.Cells[4, 2] = dt.Day + " " + a + " " + dt.Year + " - ci il";

            if (comboBox3.Text == "manat") oSheet.Cells[22, 8] = "AZN";
            else if (comboBox3.Text == "dollar") oSheet.Cells[22, 8] = "USD";
            else if (comboBox3.Text == "avro") oSheet.Cells[22, 8] = "EUR";

            reqemler();
            oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + txtmanat.Text;
            
            oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: " + txt2.Text;
            if (comboBox3.Text == "dollar") oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: " + txt2.Text.Substring(0, txt2.Text.Length - 15) + ", 00 " + "ABŞ dolları";
            oSheet.Cells[28, 2] = "D1. Ödənişin təyinatı və əsas / Payment details: " + txtteyinat.Text;
            oSheet.Cells[31, 2] = "D2. Ödənişlə əlaqədar əlavə informasiya / Narrative:";

            MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1=" + "'" + comboBox1.Text + "'");
            MyData.dtmainrekvizitler=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

            string adi = MyData.dtmainrekvizitler.Rows[0][1].ToString();
            try
            {
                for (int t = 0; t < MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; t++)
                {
                    if (MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(t, 1) == "(") { adi = MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(0, t); t = MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; }
                }
            }
            catch { }

            oSheet.Cells[17, 2] = "Adı / Name: " + adi;
            oSheet.Cells[18, 2] = "Hesab № / Acc. №: " + MyData.dtmainrekvizitler.Rows[0][2].ToString();
            oSheet.Cells[20, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0][3].ToString();

            oSheet.Cells[8, 2] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0][4].ToString();
            oSheet.Cells[9, 2] = "Kodu / Code: " + MyData.dtmainrekvizitler.Rows[0][5].ToString();
            oSheet.Cells[10, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0][6].ToString();
            oSheet.Cells[11, 2] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainrekvizitler.Rows[0][7].ToString();
            if (comboBox2.Text == "“AGLizinq” QSC (UNİ USD Daxili)") oSheet.Cells[11, 2] = "Müxbir hesab / Corr. Acc.:";
            oSheet.Cells[13, 2] = "S. W. I. F. T. Bik: " + MyData.dtmainrekvizitler.Rows[0][8].ToString();

            MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1=" + "'" + comboBox2.Text + "'");
            MyData.dtmainrekvizitler=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

            adi = MyData.dtmainrekvizitler.Rows[0][1].ToString();
            try
            {
                for (int t = 0; t < MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; t++)
                {
                    if (MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(t, 1) == "(") { adi = MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(0, t); t = MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; }
                }
            }
            catch { }

            oSheet.Cells[17, 8] = "Adı / Name: " + adi;

            oSheet.Cells[18, 8] = "Hesab № / Acc. №: " + MyData.dtmainrekvizitler.Rows[0][2].ToString();
            oSheet.Cells[20, 8] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0][3].ToString();
            if (MyData.dtmainrekvizitler.Rows[0][9].ToString() != "") oSheet.Cells[20, 8] = "Ş/V: " + MyData.dtmainrekvizitler.Rows[0][9].ToString();

            oSheet.Cells[8, 8] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0][4].ToString();
            oSheet.Cells[9, 8] = "Kodu / Code: " + MyData.dtmainrekvizitler.Rows[0][5].ToString();
            oSheet.Cells[10, 8] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0][6].ToString();
            oSheet.Cells[11, 8] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainrekvizitler.Rows[0][7].ToString();
            oSheet.Cells[13, 8] = "S. W. I. F. T. Bik: " + MyData.dtmainrekvizitler.Rows[0][8].ToString();

            MyData.updateCommand("baza.accdb", "UPDATE plateshkanomre SET nomre=" + "'" + txtnomresi.Text + "'");
            
            plateshkanomreRefresh();

            MyData.updateCommand("baza.accdb", "UPDATE plateshkateyinat SET a2=" + "'" + txtteyinat.Text + "'" + "WHERE a1=" + "'" + comboBox2.Text + "'");
            
            oXL.Visible = false;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
            oSheet.PrintOut(1, 1, 3);
            oXL.DisplayAlerts = false; 
            oWB.Close(SaveChanges: true);
            oXL.Application.Quit();

            MyData.insertCommand("baza.accdb", "insert into Emeliyyatlar (a1, a2, a3) Values ('" + DateTime.Now + "', 'ÖDƏNİŞ TAPŞIRIĞI - " + comboBox1.Text + " to " + comboBox2.Text + " - " + txtmanat.Text + " " + comboBox3.Text + " (" + txtteyinat.Text + ")','" + Environment.MachineName + "')");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (!MyCheck.davamYesNo()) return;

            try
            {
                if (txtmanat.Text.Substring(txtmanat.Text.Length - 3, 1) != ".")
                {
                    MessageBox.Show("Məbləğ düzgün qeyd olunmayıb." + Environment.NewLine + "Nümunə: 123.45"); return;
                }
            }
            catch { MessageBox.Show("Məbləğ düzgün qeyd olunmayıb." + Environment.NewLine + "Nümunə: 123.45"); return; }

            button5.Enabled = false;
            button7.Enabled = false;

            try { File.Copy("Plateshka.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Odenis Tapsirigi - " + txtmanat.Text + ".xlsx", true); }
            catch { MessageBox.Show("'Plateshka.xlsx' tapılmadı."); }

            try
            {
                MyData.selectCommand("baza.accdb", "Select * From plateshkaARXIV where №=" + "'" + txtnomresi.Text + "/" + txtnomresi2.Text + "' and HARA='" + comboBox2.Text + "'");
                MyData.dtmainArxiv=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainArxiv);

                if (MyData.dtmainArxiv.Rows.Count > 0)
                {
                    MyData.updateCommand("baza.accdb", "UPDATE plateshkaARXIV SET "
                                                                                         + "TARİX =" + "'" + dttarix.Text + "',"
                                                                                         + "№ =" + "'" + txtnomresi.Text + "/" + txtnomresi2.Text + "',"
                                                                                         + "HARDAN =" + "'" + comboBox1.Text + "',"
                                                                                         + "HARA =" + "'" + comboBox2.Text + "',"
                                                                                         + "[ÖDƏNİŞİN TƏYİNATI] =" + "'" + txtteyinat.Text + "',"
                                                                                         + "MƏBLƏĞ =" + "'" + txtmanat.Text + "',"
                                                                                         + "VALYUTA =" + "'" + comboBox3.Text + "'"
                                                                                         + " WHERE №=" + "'" + txtnomresi.Text + "/" + txtnomresi2.Text + "'");

                    
                    
                    ARXIV();
                }
                else
                {
                    MyData.insertCommand("baza.accdb", "insert into plateshkaARXIV (TARİX, №, HARDAN, HARA, [ÖDƏNİŞİN TƏYİNATI], MƏBLƏĞ, VALYUTA) Values ('" + dttarix.Text + "', '" + txtnomresi.Text + "/" + txtnomresi2.Text + "', '" + comboBox1.Text + "', '" + comboBox2.Text + "', '" + txtteyinat.Text + "', '" + txtmanat.Text + "', '" + comboBox3.Text + "')");

                    ARXIV();
                }

            }
            catch { MessageBox.Show("Arxivə kopyalamadı.."); };

            if (comboBox2.Text == "") { MessageBox.Show("Benefisiar (Alan) müştəri seçilməyib.."); return; }

            DateTime dt = dttarix.Value.Date;
            string a = MyChange.TarixSozle(dt);

            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Odenis Tapsirigi - " + txtmanat.Text + ".xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

            if (txtmanat.Text == "") txtmanat.Text = "0";

            oSheet.Cells[2, 2] = "ÖDƏNİŞ TAPŞIRIĞI № " + txtnomresi.Text + "/" + txtnomresi2.Text;
            oSheet.Cells[4, 2] = dt.Day + " " + a + " " + dt.Year + " - ci il";

            if (comboBox3.Text == "manat") oSheet.Cells[22, 8] = "AZN";
            if (comboBox3.Text == "dollar") oSheet.Cells[22, 8] = "USD";
            if (comboBox3.Text == "avro") oSheet.Cells[22, 8] = "EUR";

            reqemler();
            oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + txtmanat.Text;
            oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: " + txt2.Text;
            if (comboBox3.Text == "dollar") oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: " + txt2.Text.Substring(0, txt2.Text.Length-15) + ", 00 " + "ABŞ dolları" ;
            oSheet.Cells[28, 2] = "D1. Ödənişin təyinatı və əsas / Payment details: " + txtteyinat.Text;
            oSheet.Cells[31, 2] = "D2. Ödənişlə əlaqədar əlavə informasiya / Narrative:";

            MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1=" + "'" + comboBox1.Text + "'");
            MyData.dtmainrekvizitler=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

            string adi = MyData.dtmainrekvizitler.Rows[0][1].ToString();
            try
            {
                for (int t = 0; t < MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; t++)
                {
                    if (MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(t, 1) == "(") { adi = MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(0, t); t = MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; }
                }
            }
            catch { }

            oSheet.Cells[17, 2] = "Adı / Name: " + adi;
            oSheet.Cells[18, 2] = "Hesab № / Acc. №: " + MyData.dtmainrekvizitler.Rows[0][2].ToString();
            oSheet.Cells[20, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0][3].ToString();

            oSheet.Cells[8, 2] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0][4].ToString();
            oSheet.Cells[9, 2] = "Kodu / Code: " + MyData.dtmainrekvizitler.Rows[0][5].ToString();
            oSheet.Cells[10, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0][6].ToString();
            oSheet.Cells[11, 2] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainrekvizitler.Rows[0][7].ToString();
            if (comboBox2.Text == "“AGLizinq” QSC (UNİ USD Daxili)") oSheet.Cells[11, 2] = "Müxbir hesab / Corr. Acc.:";
            oSheet.Cells[13, 2] = "S. W. I. F. T. Bik: " + MyData.dtmainrekvizitler.Rows[0][8].ToString();
           
            MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1=" + "'" + comboBox2.Text + "'");
            MyData.dtmainrekvizitler=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

            adi= MyData.dtmainrekvizitler.Rows[0][1].ToString();
            try
            {
                for (int t = 0; t < MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; t++)
                {
                    if (MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(t, 1) == "(") { adi = MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(0, t); t = MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; }
                }
            }
            catch { }

            oSheet.Cells[17, 8] = "Adı / Name: " + adi;

            oSheet.Cells[18, 8] = "Hesab № / Acc. №: " + MyData.dtmainrekvizitler.Rows[0][2].ToString();
            oSheet.Cells[20, 8] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0][3].ToString();
            if (MyData.dtmainrekvizitler.Rows[0][9].ToString() != "") oSheet.Cells[20, 8] = "Ş/V: " + MyData.dtmainrekvizitler.Rows[0][9].ToString();

            oSheet.Cells[8, 8] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0][4].ToString();
            oSheet.Cells[9, 8] = "Kodu / Code: " + MyData.dtmainrekvizitler.Rows[0][5].ToString();
            oSheet.Cells[10, 8] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0][6].ToString();
            oSheet.Cells[11, 8] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainrekvizitler.Rows[0][7].ToString();
            oSheet.Cells[13, 8] = "S. W. I. F. T. Bik: " + MyData.dtmainrekvizitler.Rows[0][8].ToString();
          
            MyData.updateCommand("baza.accdb", "UPDATE plateshkanomre SET nomre=" + "'" + txtnomresi.Text + "'");
            
            plateshkanomreRefresh();

            MyData.updateCommand("baza.accdb", "UPDATE plateshkateyinat SET a2=" + "'" + txtteyinat.Text + "'" + "WHERE a1=" + "'" + comboBox2.Text + "'");
            
            MyData.insertCommand("baza.accdb", "insert into Emeliyyatlar (a1, a2, a3) Values ('" + DateTime.Now + "', 'ÖDƏNİŞ TAPŞIRIĞI - " + comboBox1.Text + " to " + comboBox2.Text + " - " + txtmanat.Text + " " + comboBox3.Text + " (" + txtteyinat.Text + ")','" + Environment.MachineName + "')");
           
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (txtedv1.Text == "") { MessageBox.Show("Müştərinin Adı qeyd olunmayıb"); return; }
            if (txtedv2.Text == "") { MessageBox.Show("Müştərinin Hesabı qeyd olunmayıb"); return; }
            if (txtedv22.Text == "") { MessageBox.Show("Müştərinin Hesabı qeyd olunmayıb"); return; }
            if (txtedv3.Text == "") { MessageBox.Show("Müştərinin VÖEN-i qeyd olunmayıb"); return; }
            if (txtedv4.Text == "") { MessageBox.Show("Bankın Adı qeyd olunmayıb"); return; }
            if (txtedv5.Text == "") { MessageBox.Show("Bankın Kodu qeyd olunmayıb"); return; }
            if (txtedv6.Text == "") { MessageBox.Show("Bankın VÖEN-i qeyd olunmayıb"); return; }
            if (txtedv7.Text == "") { MessageBox.Show("Bankın Müxbir Hesabı qeyd olunmayıb"); return; }
            if (txtedv77.Text == "") { MessageBox.Show("Bankın Müxbir Hesabı qeyd olunmayıb"); return; }
            if (txtedv8.Text == "") { MessageBox.Show("Bankın S.W.I.F.T-i qeyd olunmayıb"); return; } 

            int k, k2 = 0;

            MyData.selectCommand("baza.accdb", "Select * From plateshkaEDV");
            MyData.dtmainedv=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainedv);

            for (k = 0; k < MyData.dtmainedv.Rows.Count; k++)
            {
                if (txtedv1.Text == MyData.dtmainedv.Rows[k]["a1"].ToString()) k2 += 1;
            }

            if (k2 > 0) { MessageBox.Show("Qeyd olunan müştəri artıq mövcuddur."); return; }

            MyData.insertCommand("baza.accdb", "insert into plateshkaEDV (a1, a2, a3, a4, a5, a6, a7, a8, a9, a10) Values ('" + txtedv1.Text + "','" + txtedv2.Text + " " + txtedv22.Text + "','" + txtedv3.Text + "','" + txtedv4.Text + "','" + txtedv5.Text + "','" + txtedv6.Text + "','" + txtedv7.Text + " " + txtedv77.Text + "','" + txtedv8.Text + "','" + txtTesnifatKod.Text + "','" + txtSeviyyeKod.Text + "')");
            
            EDVRefresh();

            MyData.selectCommand("baza.accdb", "Select * From plateshkateyinat where a1 Like '%" + txtedv1.Text + "%'");
            MyData.dtmainTeyinat=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainTeyinat);

            if (MyData.dtmainTeyinat.Rows.Count == 0)
            {
                MyData.insertCommand("baza.accdb", "insert into plateshkateyinat (a1, a2) Values ('" + txtedv1.Text + "','" + txtteyinatEDV.Text + "')");
                
            }
            else
            {
                MyData.updateCommand("baza.accdb", "UPDATE plateshkateyinat SET a2='" + txtteyinatEDV.Text + "' WHERE a1 Like '%" + txtedv1.Text + "%'");
                 
            }

            MessageBox.Show("Müştəri əlavə olundu.");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (!MyCheck.davamYesNo()) return;

            try
            {
                if (txtedvmanat.Text.Substring(txtedvmanat.Text.Length - 3, 1) != ".")
                {
                    MessageBox.Show("Məbləğ düzgün qeyd olunmayıb." + Environment.NewLine + "Nümunə: 123.45"); return;
                }
            }
            catch { MessageBox.Show("Məbləğ düzgün qeyd olunmayıb." + Environment.NewLine + "Nümunə: 123.45"); return; }

            button6.Enabled = false;
            button8.Enabled = false; 
            
            try { File.Copy("EDV.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Odenis Tapsirigi - " + txtedvmanat.Text +  ".xlsx", true); }
            catch { MessageBox.Show("'EDV.xlsx' tapılmadı."); }

            try
            {
                MyData.selectCommand("baza.accdb", "Select * From plateshkaARXIV where №='" + txtedvnomre.Text + "/" + txtedvnomre2.Text + "' and HARA='" + cbedv2.Text + "'");
                MyData.dtmainArxiv=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainArxiv);

                if (MyData.dtmainArxiv.Rows.Count > 0)
                {
                    MyData.updateCommand("baza.accdb", "UPDATE plateshkaARXIV SET "
                                                                                         + "TARİX =" + "'" + dtedvtarix.Text + "',"
                                                                                         + "№ =" + "'" + txtedvnomre.Text + "/" + txtedvnomre2.Text + "',"
                                                                                         + "HARDAN =" + "'" + cbedv1.Text + "',"
                                                                                         + "HARA =" + "'" + cbedv2.Text + "',"
                                                                                         + "[ÖDƏNİŞİN TƏYİNATI] =" + "'" + txtedvteyinat.Text + "',"
                                                                                         + "MƏBLƏĞ =" + "'" + txtedvmanat.Text +  "',"
                                                                                         + "VALYUTA =" + "'" + comboBox6.Text + "'"
                                                                                         + " WHERE №=" + "'" + txtedvnomre.Text + "/" + txtedvnomre2.Text + "'");

                    
                    
                    ARXIV();
                }
                else
                {
                    MyData.insertCommand("baza.accdb", "insert into plateshkaARXIV (TARİX, №, HARDAN, HARA, [ÖDƏNİŞİN TƏYİNATI], MƏBLƏĞ, VALYUTA) Values ('" + dtedvtarix.Text + "', '" + txtedvnomre.Text + "/" + txtedvnomre2.Text + "', '" + cbedv1.Text + "', '" + cbedv2.Text + "', '" + txtedvteyinat.Text + "', '" + txtedvmanat.Text +  "', '" + comboBox6.Text + "')");
                    
                    ARXIV();
                }

            }
            catch { MessageBox.Show("Arxivə kopyalamadı.."); };


            if (cbedv2.Text == "") { MessageBox.Show("Benefisiar (Alan) müştəri seçilməyib.."); return; }

            DateTime dt = dtedvtarix.Value.Date;
            string a = MyChange.TarixSozle(dt);

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Odenis Tapsirigi - " + txtedvmanat.Text +  ".xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            if (txtedvmanat.Text == "") txtmanat.Text = "0";

            oSheet.Cells[3, 3] = "'A' formatlı ödəniş tapşırığı № " + txtedvnomre.Text + "/" + txtedvnomre2.Text;
            oSheet.Cells[4, 3] = dt.Day+ " " + a + " " + dt.Year + " - ci il";

            if (comboBox6.Text == "manat") oSheet.Cells[22, 9] = "AZN";
            if (comboBox6.Text == "dollar") oSheet.Cells[22, 9] = "USD";
            if (comboBox6.Text == "avro") oSheet.Cells[22, 9] = "EUR";

            reqemler2();
            oSheet.Cells[24, 3] = "Məbləğ rəqəmlə: " + txtedvmanat.Text;
            oSheet.Cells[25, 3] = "Məbləğ yazı ilə / İn words: " + txtedvherfle.Text; 
            if (comboBox3.Text == "dollar") oSheet.Cells[25, 3] = "Məbləğ yazı ilə / İn words: " + txtedvherfle.Text.Substring(0, txtedvherfle.Text.Length - 15) + ", 00 " + "ABŞ dolları";
            oSheet.Cells[26, 3] = "D1. Ödənişin təyinatı və əsas / Payment details: " + txtedvteyinat.Text;

            MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1=" + "'" + cbedv1.Text + "'");
            MyData.dtmainrekvizitler=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

            string adi = MyData.dtmainrekvizitler.Rows[0][1].ToString();
            try
            {
                for (int t = 0; t < MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; t++)
                {
                    if (MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(t, 1) == "(") { adi = MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(0, t); t = MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; }
                }
            }
            catch { }
            oSheet.Cells[17, 3] = "Adı / Name: " + adi;

            oSheet.Cells[18, 3] = "Hesab № / Acc. №: " + MyData.dtmainrekvizitler.Rows[0][2].ToString();
            oSheet.Cells[20, 3] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0][3].ToString();

            oSheet.Cells[8, 3] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0][4].ToString();
            oSheet.Cells[9, 3] = "Kodu / Code: " + MyData.dtmainrekvizitler.Rows[0][5].ToString();
            oSheet.Cells[10, 3] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0][6].ToString();
            oSheet.Cells[11, 3] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainrekvizitler.Rows[0][7].ToString();
            oSheet.Cells[13, 3] = "S. W. I. F. T. Bik: " + MyData.dtmainrekvizitler.Rows[0][8].ToString();
            
            MyData.selectCommand("baza.accdb", "Select * From plateshkaEDV WHERE a1=" + "'" + cbedv2.Text + "'");
            MyData.dtmainedv=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainedv);

            adi = MyData.dtmainedv.Rows[0]["a1"].ToString();
            try
            {
                for (int t = 0; t < MyData.dtmainedv.Rows[0]["a1"].ToString().Length; t++)
                {
                    if (MyData.dtmainedv.Rows[0]["a1"].ToString().Substring(t, 1) == "(") { adi = MyData.dtmainedv.Rows[0]["a1"].ToString().Substring(0, t); t = MyData.dtmainedv.Rows[0]["a1"].ToString().Length; }
                }
            }
            catch { }
            oSheet.Cells[17, 9] = "Adı / Name: " + adi;

            oSheet.Cells[18, 9] = "Hesab № / Acc. №: " + MyData.dtmainedv.Rows[0]["a2"].ToString();
            oSheet.Cells[20, 9] = "VÖEN / Tax İD: " + MyData.dtmainedv.Rows[0]["a3"].ToString();

            oSheet.Cells[8, 9] = "Adı / Name: " + MyData.dtmainedv.Rows[0]["a4"].ToString();
            oSheet.Cells[9, 9] = "Kodu / Code: " + MyData.dtmainedv.Rows[0]["a5"].ToString();
            oSheet.Cells[10, 9] = "VÖEN / Tax İD: " + MyData.dtmainedv.Rows[0]["a6"].ToString();
            oSheet.Cells[11, 9] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainedv.Rows[0]["a7"].ToString();
            oSheet.Cells[13, 9] = "S. W. I. F. T. Bik: " + MyData.dtmainedv.Rows[0]["a8"].ToString();
            oSheet.Cells[31, "F"] = MyData.dtmainedv.Rows[0]["a9"].ToString();
            oSheet.Cells[31, "L"] = MyData.dtmainedv.Rows[0]["a10"].ToString();
            
            MyData.updateCommand("baza.accdb", "UPDATE plateshkaEDVnomre SET nomre=" + "'" + txtedvnomre.Text + "'");
            
            EDVnomreRefresh();

            //-----------------------teyinatlarin yenilenmesi
           try
            {
                MyData.updateCommand("baza.accdb", "UPDATE plateshkateyinat SET a2='" + txtedvteyinat.Text + "' WHERE a1 Like '%" + cbedv2.Text + "%'");
                
            }
            catch { MessageBox.Show("Teyinat Yenilenmedi."); }


            oXL.Visible = false;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
            oSheet.PrintOut(1, 1, 3); 
            oXL.DisplayAlerts = false; 
            oWB.Close(SaveChanges: true);
            oXL.Application.Quit();

            MyData.insertCommand("baza.accdb", "insert into Emeliyyatlar (a1, a2, a3) Values ('" + DateTime.Now + "', 'ÖDƏNİŞ TAPŞIRIĞI - " + cbedv2.Text + " - " + txtedvmanat.Text +  " " + comboBox6.Text + " - ƏDV (" + txtedvteyinat.Text + ")','" + Environment.MachineName + "')");
            
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (txtedvnomre.Enabled == false || txtedvnomre2.Enabled == false) { txtedvnomre.Enabled = true; txtedvnomre2.Enabled = true; return; }
            if (txtedvnomre.Enabled == true || txtedvnomre2.Enabled == true) { txtedvnomre.Enabled = false; txtedvnomre2.Enabled = false; }

        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (dtedvtarix.Enabled == false) { dtedvtarix.Enabled = true; return; }
            if (dtedvtarix.Enabled == true) { dtedvtarix.Enabled = false; }

        }

        private void cbedv2_SelectedIndexChanged(object sender, EventArgs e)
        {
            button6.Enabled = true;
            button8.Enabled = true; 
            
            txtedvmanat.Text = "0"; 

            MyData.selectCommand("baza.accdb", "Select * From plateshkateyinat WHERE a1=" + "'" + cbedv2.Text + "'");
            MyData.dtmainTeyinat=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainTeyinat);
            try
            {
                txtedvteyinat.Text = MyData.dtmainTeyinat.Rows[0][1].ToString();
            }
            catch { MessageBox.Show("TEYINATDA SEHV"); };

            try { if (cbedv2.Text.Substring(1, 8) == "Azpetrol") txtedvmanat.Text = "152.54"; }
            catch { }
            try { if (cbedv2.Text.Substring(1, 8) == "Progress") txtedvmanat.Text = "45.76"; }
            catch { }

            MyData.selectCommand("baza.accdb", "Select * From PlateshkaEDV WHERE a1 Like " + "'%" + cbedv2.Text + "%'");
            MyData.dtmainedv=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainedv);

            lbbankadi4.Text = "Bank - " + MyData.dtmainedv.Rows[0]["a4"].ToString() + Environment.NewLine + "H/h - " + MyData.dtmainedv.Rows[0]["a2"].ToString() + Environment.NewLine + "M/h - " + MyData.dtmainedv.Rows[0]["a7"].ToString();

            //SON ODENISLER EDV UCUN
            try
            {
                MyData.selectCommand("baza.accdb", "Select * From plateshkaARXIV WHERE HARA Like '%" + cbedv2.Text + "%'");
                MyData.dtmainArxiv = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainArxiv);

                btSonXeberEDV.Text = "SON ÖDƏNİŞLƏR:" + Environment.NewLine;
                try { btSonXeberEDV.Text += MyData.dtmainArxiv.Rows[0]["HARA"].ToString() + " ♦ ♦ ♦ " + MyData.dtmainArxiv.Rows[0]["TARİX"].ToString() + " ♦ ♦ " + MyData.dtmainArxiv.Rows[0]["MƏBLƏĞ"].ToString() + " AZN" + Environment.NewLine; }
                catch { }
                try { btSonXeberEDV.Text += MyData.dtmainArxiv.Rows[1]["HARA"].ToString() + " ♦ ♦ ♦ " + MyData.dtmainArxiv.Rows[1]["TARİX"].ToString() + " ♦ ♦ " + MyData.dtmainArxiv.Rows[1]["MƏBLƏĞ"].ToString() + " AZN" + Environment.NewLine; }
                catch { }
                try { btSonXeberEDV.Text += MyData.dtmainArxiv.Rows[2]["HARA"].ToString() + " ♦ ♦ ♦ " + MyData.dtmainArxiv.Rows[2]["TARİX"].ToString() + " ♦ ♦ " + MyData.dtmainArxiv.Rows[2]["MƏBLƏĞ"].ToString() + " AZN" + Environment.NewLine; }
                catch { }
                try { btSonXeberEDV.Text += MyData.dtmainArxiv.Rows[3]["HARA"].ToString() + " ♦ ♦ ♦ " + MyData.dtmainArxiv.Rows[3]["TARİX"].ToString() + " ♦ ♦ " + MyData.dtmainArxiv.Rows[3]["MƏBLƏĞ"].ToString() + " AZN" + Environment.NewLine; }
                catch { }
            }
            catch { };
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (!MyCheck.davamYesNo()) return;

            try
            {
                if (txtedvmanat.Text.Substring(txtedvmanat.Text.Length - 3, 1) != ".")
                {
                    MessageBox.Show("Məbləğ düzgün qeyd olunmayıb." + Environment.NewLine + "Nümunə: 123.45"); return;
                }
            }
            catch { MessageBox.Show("Məbləğ düzgün qeyd olunmayıb." + Environment.NewLine + "Nümunə: 123.45"); return; }

            button6.Enabled = false;
            button8.Enabled = false; 
            
            try { File.Copy("EDV.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Odenis Tapsirigi - " + txtedvmanat.Text +  ".xlsx", true); }
            catch { MessageBox.Show("'EDV.xlsx' tapılmadı."); }

            try
            {
                
                
                
                MyData.selectCommand("baza.accdb", "Select * From plateshkaARXIV where №=" + "'" + txtedvnomre.Text + "/" + txtedvnomre2.Text + "' and HARA='" + cbedv2.Text + "'");
                MyData.dtmainArxiv=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainArxiv);

                if (MyData.dtmainArxiv.Rows.Count > 0)
                {
                    MyData.updateCommand("baza.accdb", "UPDATE plateshkaARXIV SET "
                                                                                         + "TARİX =" + "'" + dtedvtarix.Text + "',"
                                                                                         + "№ =" + "'" + txtedvnomre.Text + "/" + txtedvnomre2.Text + "',"
                                                                                         + "HARDAN =" + "'" + cbedv1.Text + "',"
                                                                                         + "HARA =" + "'" + cbedv2.Text + "',"
                                                                                         + "[ÖDƏNİŞİN TƏYİNATI] =" + "'" + txtedvteyinat.Text + "',"
                                                                                         + "MƏBLƏĞ =" + "'" + txtedvmanat.Text +  "',"
                                                                                         + "VALYUTA =" + "'" + comboBox6.Text + "'"
                                                                                         + " WHERE №=" + "'" + txtedvnomre.Text + "/" + txtedvnomre2.Text + "'");

                    
                    
                    ARXIV();
                }
                else
                {
                    MyData.insertCommand("baza.accdb", "insert into plateshkaARXIV (TARİX, №, HARDAN, HARA, [ÖDƏNİŞİN TƏYİNATI], MƏBLƏĞ, VALYUTA) Values ('" + dtedvtarix.Text + "', '" + txtedvnomre.Text + "/" + txtedvnomre2.Text + "', '" + cbedv1.Text + "', '" + cbedv2.Text + "', '" + txtedvteyinat.Text + "', '" + txtedvmanat.Text +  "', '" + comboBox6.Text + "')");
                    
                    ARXIV();
                }

            }
            catch { MessageBox.Show("Arxivə kopyalamadı.."); };


            if (cbedv2.Text == "") { MessageBox.Show("Benefisiar (Alan) müştəri seçilməyib.."); return; }

            DateTime dt = dtedvtarix.Value.Date;
            string a = MyChange.TarixSozle(dt);

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Odenis Tapsirigi - " + txtedvmanat.Text +  ".xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

            if (txtedvmanat.Text == "") txtmanat.Text = "0";

            oSheet.Cells[3, 3] = "'A' formatlı ödəniş tapşırığı № " + txtedvnomre.Text + "/" + txtedvnomre2.Text;
            oSheet.Cells[4, 3] = dt.Day + " " + a + " " + dt.Year + " - ci il";

            if (comboBox6.Text == "manat") oSheet.Cells[22, 9] = "AZN";
            if (comboBox6.Text == "dollar") oSheet.Cells[22, 9] = "USD";
            if (comboBox6.Text == "avro") oSheet.Cells[22, 9] = "EUR";

            reqemler2();
            oSheet.Cells[24, 3] = "Məbləğ rəqəmlə: " + txtedvmanat.Text;
            
            oSheet.Cells[25, 3] = "Məbləğ yazı ilə / İn words: " + txtedvherfle.Text;
            if (comboBox3.Text == "dollar") oSheet.Cells[25, 3] = "Məbləğ yazı ilə / İn words: " + txtedvherfle.Text.Substring(0, txtedvherfle.Text.Length - 15) + ", 00 " + "ABŞ dolları";
            oSheet.Cells[26, 3] = "D1. Ödənişin təyinatı və əsas / Payment details: " + txtedvteyinat.Text;

            MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1=" + "'" + cbedv1.Text + "'");
            MyData.dtmainrekvizitler=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

            string adi = MyData.dtmainrekvizitler.Rows[0][1].ToString();
            try
            {
                for (int t = 0; t < MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; t++)
                {
                    if (MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(t, 1) == "(") { adi = MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(0, t); t = MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; }
                }
            }
            catch { }
            oSheet.Cells[17, 3] = "Adı / Name: " + adi;

            oSheet.Cells[18, 3] = "Hesab № / Acc. №: " + MyData.dtmainrekvizitler.Rows[0][2].ToString();
            oSheet.Cells[20, 3] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0][3].ToString();

            oSheet.Cells[8, 3] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0][4].ToString();
            oSheet.Cells[9, 3] = "Kodu / Code: " + MyData.dtmainrekvizitler.Rows[0][5].ToString();
            oSheet.Cells[10, 3] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0][6].ToString();
            oSheet.Cells[11, 3] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainrekvizitler.Rows[0][7].ToString();
            oSheet.Cells[13, 3] = "S. W. I. F. T. Bik: " + MyData.dtmainrekvizitler.Rows[0][8].ToString();
            //////---------------------------------------------------------------------------------------

            ///////-------------------------Benefisiar alan bank ve musterinin rekvizitleri ucun
            
            
            MyData.selectCommand("baza.accdb", "Select * From plateshkaEDV WHERE a1=" + "'" + cbedv2.Text + "'");
            MyData.dtmainedv=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainedv);

            adi = MyData.dtmainedv.Rows[0]["a1"].ToString();
            try
            {
                for (int t = 0; t < MyData.dtmainedv.Rows[0]["a1"].ToString().Length; t++)
                {
                    if (MyData.dtmainedv.Rows[0]["a1"].ToString().Substring(t, 1) == "(") { adi = MyData.dtmainedv.Rows[0]["a1"].ToString().Substring(0, t); t = MyData.dtmainedv.Rows[0]["a1"].ToString().Length; }
                }
            }
            catch { }
            oSheet.Cells[17, 9] = "Adı / Name: " + adi;

            oSheet.Cells[18, 9] = "Hesab № / Acc. №: " + MyData.dtmainedv.Rows[0]["a2"].ToString();
            oSheet.Cells[20, 9] = "VÖEN / Tax İD: " + MyData.dtmainedv.Rows[0]["a3"].ToString();
            oSheet.Cells[8, 9] = "Adı / Name: " + MyData.dtmainedv.Rows[0]["a4"].ToString();
            oSheet.Cells[9, 9] = "Kodu / Code: " + MyData.dtmainedv.Rows[0]["a5"].ToString();
            oSheet.Cells[10, 9] = "VÖEN / Tax İD: " + MyData.dtmainedv.Rows[0]["a6"].ToString();
            oSheet.Cells[11, 9] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainedv.Rows[0]["a7"].ToString();
            oSheet.Cells[13, 9] = "S. W. I. F. T. Bik: " + MyData.dtmainedv.Rows[0]["a8"].ToString();
            oSheet.Cells[31, "F"] = MyData.dtmainedv.Rows[0]["a9"].ToString();
            oSheet.Cells[31, "L"] = MyData.dtmainedv.Rows[0]["a10"].ToString();

            MyData.updateCommand("baza.accdb", "UPDATE plateshkaEDVnomre SET nomre=" + "'" + txtedvnomre.Text + "'");
            
            EDVnomreRefresh();

            try
            {
                MyData.updateCommand("baza.accdb", "UPDATE plateshkateyinat SET a2='" + txtedvteyinat.Text + "' WHERE a1 Like '%" + cbedv2.Text + "%'");
            }
            catch { MessageBox.Show("Teyinat Yenilenmedi."); }

            MyData.insertCommand("baza.accdb", "insert into Emeliyyatlar (a1, a2, a3) Values ('" + DateTime.Now + "', 'ÖDƏNİŞ TAPŞIRIĞI - " + cbedv2.Text + " - " + txtedvmanat.Text +  " " + comboBox6.Text + " - ƏDV (" + txtedvteyinat.Text + ")','" + Environment.MachineName + "')");
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if (dateTimePicker1.Enabled == false) { dateTimePicker1.Enabled = true; return; }
            if (dateTimePicker1.Enabled == true) { dateTimePicker1.Enabled = false; }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            if (!MyCheck.davamYesNo()) return;

            try
            {
                if (txtkonvertmanat.Text.Substring(txtkonvertmanat.Text.Length - 3, 1) != ".")
                {
                    MessageBox.Show("Məbləğ düzgün qeyd olunmayıb." + Environment.NewLine + "Nümunə: 123.45"); return;
                }
            }
            catch { MessageBox.Show("Məbləğ düzgün qeyd olunmayıb." + Environment.NewLine + "Nümunə: 123.45"); return; }

            try { File.Copy("Konvert.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Konvert.xlsx", true); }
            catch { MessageBox.Show("'Konvert.xlsx' tapılmadı."); }

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Konvert.xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

            DateTime dt = dateTimePicker1.Value.Date;
            string a = MyChange.TarixSozle(dt);

            oSheet.Cells[4, 9] = dt.Day + " " + a + " " + dt.Year + " - ci il";
            oSheet.Cells[2, 1] = lbBasliq.Text;
            oSheet.Cells[7, 1] = "Alınan valyutanın adı __" + cbAlinanValyutaHesab.Text.Substring(cbAlinanValyutaHesab.Text.Length - 3, 3) + "___ hesabın nömrəsi ______" + cbAlinanValyutaHesab.Text.Substring(0, cbAlinanValyutaHesab.Text.Length - 3) + "____";
            oSheet.Cells[8, 1] = "Satılan valyutanın adı __" + cbSatilanValyutaHesab.Text.Substring(cbSatilanValyutaHesab.Text.Length - 3, 3) + "___ hesabın nömrəsi ______" + cbSatilanValyutaHesab.Text.Substring(0, cbSatilanValyutaHesab.Text.Length - 3) + "____";
            
            
            if (cbXariciValyutadaMebleg.Text == "EUR") oSheet.Cells[9, 4] = txtkonvertmanat.Text + " (" + txtherf5.Text + ") AVRO";
            else oSheet.Cells[9, 4] = txtkonvertmanat.Text + " (" + txtherf5.Text + ") ABŞ dolları";

            oSheet.Cells[11, 1] = "Alınan valyuta aşağıdakı məqsədlərə istifadə ediləcəkdir: " + textBox5.Text;

            oWB.Save();
            oXL.DisplayAlerts = false; 
            
        }

        private void button12_Click(object sender, EventArgs e)
        {
            if (!MyCheck.davamYesNo()) return;

            try
            {
                if (txtkonvertmanat.Text.Substring(txtkonvertmanat.Text.Length - 3, 1) != ".")
                {
                    MessageBox.Show("Məbləğ düzgün qeyd olunmayıb." + Environment.NewLine + "Nümunə: 123.45"); return;
                }
            }
            catch { MessageBox.Show("Məbləğ düzgün qeyd olunmayıb." + Environment.NewLine + "Nümunə: 123.45"); return; }

            try { File.Copy("Konvert.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Konvert.xlsx", true); }
            catch { MessageBox.Show("'Konvert.xlsx' tapılmadı."); }

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open(Application.StartupPath + "\\Konvert.xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

            DateTime dt = dateTimePicker1.Value.Date;
            string a = MyChange.TarixSozle(dt);

            oSheet.Cells[4, 9] = dt.Day+ " " + a + " " + dt.Year + " - ci il";
            oSheet.Cells[2, 1] = lbBasliq.Text;
            oSheet.Cells[7, 1] = "Alınan valyutanın adı __" + cbAlinanValyutaHesab.Text.Substring(cbAlinanValyutaHesab.Text.Length - 3, 3) + "___ hesabın nömrəsi ______" + cbAlinanValyutaHesab.Text.Substring(0, cbAlinanValyutaHesab.Text.Length - 3) + "____";
            oSheet.Cells[8, 1] = "Satılan valyutanın adı __" + cbSatilanValyutaHesab.Text.Substring(cbSatilanValyutaHesab.Text.Length - 3, 3) + "___ hesabın nömrəsi ______" + cbSatilanValyutaHesab.Text.Substring(0, cbSatilanValyutaHesab.Text.Length - 3) + "____";
            
            if (cbXariciValyutadaMebleg.Text == "EUR") oSheet.Cells[9, 4] = txtkonvertmanat.Text + " (" + txtherf5.Text + ") AVRO";
            else oSheet.Cells[9, 4] = txtkonvertmanat.Text + " (" + txtherf5.Text + ") ABŞ dolları";
            
            oSheet.Cells[11, 1] = "Alınan valyuta aşağıdakı məqsədlərə istifadə ediləcəkdir: " + textBox5.Text;

            oXL.Visible = false;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
            oSheet.PrintOut();
            
            oWB.Close(SaveChanges: true);
            oXL.DisplayAlerts = false; 
            oXL.Application.Quit();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            button5.Enabled = true;
            button7.Enabled = true;

            label41.Visible = false; comboBox3.Text = "manat"; txtmanat.Text = "0";
            if (comboBox1.Text == "“AGLizinq” QSC (AGB AZN)") { label41.Visible = true; label41.Text = "AGLizinq AGB AZN hesabı"; comboBox3.Text = "manat"; txtmanat.Text = "0"; }
            else if (comboBox1.Text == "“AGLizinq” QSC (AGB USD)") { label41.Visible = true; label41.Text = "AGLizinq AGB USD hesabı"; comboBox3.Text = "dollar"; txtmanat.Text = "0"; }
            else if (comboBox1.Text == "“AGLizinq” QSC (ABB USD)") { label41.Visible = true; label41.Text = "AGLizinq ABB USD hesabı"; comboBox3.Text = "dollar"; txtmanat.Text = "0"; }
            else if (comboBox1.Text == "“AGLizinq” QSC (AGB CARD USD)") { label41.Visible = true; label41.Text = "AGLizinq AGB CARD USD hesabı"; comboBox3.Text = "dollar"; txtmanat.Text = "0"; }
            else if (comboBox1.Text == "“AGLizinq” QSC (ABB AZN)") { label41.Visible = true; label41.Text = "AGLizinq ABB AZN hesabı"; comboBox3.Text = "manat"; txtmanat.Text = "0"; }
            else if (comboBox1.Text == "“AGLizinq” QSC (UNİ USD)") { label41.Visible = true; label41.Text = "AGLizinq UNİ USD hesabı"; comboBox3.Text = "dollar"; txtmanat.Text = "0"; }
            else if (comboBox1.Text == "“AGLizinq” QSC (UNİ USD Daxili)") { label41.Visible = true; label41.Text = "AGLizinq UNİ USD Daxili hesabı"; comboBox3.Text = "dollar"; txtmanat.Text = "0"; }

            MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1 Like " + "'%" + comboBox1.Text + "%'");
            MyData.dtmainrekvizitler=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

            lbbankadi1.Text = "Bank - " + MyData.dtmainrekvizitler.Rows[0]["a4"].ToString() + Environment.NewLine + "H/h - " + MyData.dtmainrekvizitler.Rows[0]["a2"].ToString() + Environment.NewLine + "M/h - " + MyData.dtmainrekvizitler.Rows[0]["a7"].ToString();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            button5.Enabled = true;
            button7.Enabled = true;

            txtmanat.Text = "0";
            txtteyinat.Text = "";  //son odenisleri her defe temizlemek ucun

            label42.Visible = false; comboBox3.Text = "manat"; txtmanat.Text = "0";
            if (comboBox2.Text == "“AGLizinq” QSC (AGB AZN)") { label42.Visible = true; label42.Text = "AGLizinq AGB AZN hesabı"; comboBox3.Text = "manat"; txtmanat.Text = "0"; }
            else if (comboBox2.Text == "“AGLizinq” QSC (AGB USD)") { label42.Visible = true; label42.Text = "AGLizinq AGB USD hesabı"; comboBox3.Text = "dollar"; txtmanat.Text = "0"; }
            else if (comboBox2.Text == "“AGLizinq” QSC (ABB USD)") { label42.Visible = true; label42.Text = "AGLizinq ABB USD hesabı"; comboBox3.Text = "dollar"; txtmanat.Text = "0"; }
            else if (comboBox2.Text == "“AGLizinq” QSC (AGB CARD USD)") { label42.Visible = true; label42.Text = "AGLizinq AGB CARD USD hesabı"; comboBox3.Text = "dollar"; txtmanat.Text = "0"; }
            else if (comboBox2.Text == "“AGLizinq” QSC (ABB AZN)") { label42.Visible = true; label42.Text = "AGLizinq ABB AZN hesabı"; comboBox3.Text = "manat"; txtmanat.Text = "0"; }
            else if (comboBox2.Text == "“AGLizinq” QSC (UNİ USD)") { label42.Visible = true; label42.Text = "AGLizinq UNİ USD hesabı"; comboBox3.Text = "dollar"; txtmanat.Text = "0"; }
            else if (comboBox2.Text == "“AGLizinq” QSC (UNİ USD Daxili)") { label42.Visible = true; label42.Text = "AGLizinq UNİ USD Daxili hesabı"; comboBox3.Text = "dollar"; txtmanat.Text = "0"; }

            MyData.selectCommand("baza.accdb", "Select * From plateshkateyinat WHERE a1=" + "'" + comboBox2.Text + "'");
            MyData.dtmainTeyinat=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainTeyinat);
            try
            { txtteyinat.Text = MyData.dtmainTeyinat.Rows[0][1].ToString(); }
            catch { MessageBox.Show("TEYINATDA SEHV"); };

            try { if (comboBox2.Text.Substring(1, 8) == "Azpetrol") { txtmanat.Text = "847.46"; } }
            catch { };
            try { if (comboBox2.Text.Substring(1, 5) == "Libra") { txtmanat.Text = "500.00"; } }
            catch { };
            try { if (comboBox2.Text.Substring(1, 8) == "Progress") { txtmanat.Text = "254.24"; } }
            catch { };

            MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1 Like '%" + comboBox2.Text + "%'");
            MyData.dtmainrekvizitler=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

            lbbankadi2.Text = "Bank - " + MyData.dtmainrekvizitler.Rows[0]["a4"].ToString() + Environment.NewLine + "H/h - " + MyData.dtmainrekvizitler.Rows[0]["a2"].ToString() + Environment.NewLine + "M/h - " + MyData.dtmainrekvizitler.Rows[0]["a7"].ToString();

            try
            {
                MyData.selectCommand("baza.accdb", "Select * From plateshkaARXIV WHERE HARA Like '%" + comboBox2.Text + "%'");
                MyData.dtmainArxiv = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainArxiv);

                button23.Text = "SON ÖDƏNİŞLƏR:" + Environment.NewLine;
                try { button23.Text += MyData.dtmainArxiv.Rows[0]["HARA"].ToString() + " ♦ ♦ ♦ " + MyData.dtmainArxiv.Rows[0]["TARİX"].ToString() + " ♦ ♦ " + MyData.dtmainArxiv.Rows[0]["MƏBLƏĞ"].ToString() + " AZN" + Environment.NewLine; }
                catch { }
                try { button23.Text += MyData.dtmainArxiv.Rows[1]["HARA"].ToString() + " ♦ ♦ ♦ " + MyData.dtmainArxiv.Rows[1]["TARİX"].ToString() + " ♦ ♦ " + MyData.dtmainArxiv.Rows[1]["MƏBLƏĞ"].ToString() + " AZN" + Environment.NewLine; }
                catch { }
                try { button23.Text += MyData.dtmainArxiv.Rows[2]["HARA"].ToString() + " ♦ ♦ ♦ " + MyData.dtmainArxiv.Rows[2]["TARİX"].ToString() + " ♦ ♦ " + MyData.dtmainArxiv.Rows[2]["MƏBLƏĞ"].ToString() + " AZN" + Environment.NewLine; }
                catch { }
                try { button23.Text += MyData.dtmainArxiv.Rows[3]["HARA"].ToString() + " ♦ ♦ ♦ " + MyData.dtmainArxiv.Rows[3]["TARİX"].ToString() + " ♦ ♦ " + MyData.dtmainArxiv.Rows[3]["MƏBLƏĞ"].ToString() + " AZN" + Environment.NewLine; }
                catch { }
            }
            catch { };

        }

        private void textBox4_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try { ARXIV(); } catch { }
            }
        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                MyData.deleteCommand("baza.accdb", "DELETE FROM plateshkaARXIV WHERE Код Like '%" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Код"].Value.ToString() + "%'");
                
                MessageBox.Show("Tapşırıq yerinə yetirildi.");

                ARXIV();
            }
            catch { };
           
        }

        private void bttemizle_Click(object sender, EventArgs e)
        {
            if (!MyCheck.davamYesNo()) return;

            txtbankadi.Text = "";
            txtbankkodu.Text = "";
            txtbankvoen.Text = "";
            txtmuxhesab1.Text = "";
            txtmuxhesab2.Text = "";
            txtsvift.Text = "";
            txtmusteriadi.Text = "";
            txtmusterivoen.Text = "";
            txthesab1.Text = "";
            txthesab2.Text = "";
            txtsexsiyyet.Text = "";
            txtesas.Text = "";
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (!MyCheck.davamYesNo()) return;

            txtedv1.Text = "";
            txtedv2.Text = "";
            txtedv3.Text = "";
            txtedv4.Text = "";
            txtedv5.Text = "";
            txtedv6.Text = "";
            txtedv7.Text = "";
            txtedv77.Text = "";
            txtedv8.Text = "";
            txtedv1.Text = "";
            txtedv2.Text = "";
            txtedv22.Text = "";
            txtedv3.Text = "";
            txtSeviyyeKod.Text = "";
            txtTesnifatKod.Text = "";
            txtteyinatEDV.Text = "";
        }

        private void label1_Click(object sender, EventArgs e)
        {
            if (comboBox1.Enabled == false) { comboBox1.Enabled = true; return; }
            comboBox1.Enabled = false; ;

        }

        private void label34_Click(object sender, EventArgs e)
        {
            if (cbedv1.Enabled == false) { cbedv1.Enabled = true; return; }
            cbedv1.Enabled = false; ;
        }

        private void txtteyinat_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                MyData.selectCommand("baza.accdb", "Select * From plateshkateyinat WHERE a2 Like " + "'%" + txtteyinat.Text + "%'");
                MyData.dtmainTeyinat=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainTeyinat);
                try
                { txtteyinat.Text = MyData.dtmainTeyinat.Rows[0][1].ToString(); }
                catch { };

            }
        }

        private void txtbankadi_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a4 Like " + "'%" + txtbankadi.Text + "%'");
                MyData.dtmainrekvizitler=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);
                try
                { 
                    txtbankadi.Text = MyData.dtmainrekvizitler.Rows[0][4].ToString();
                    txtbankkodu.Text = MyData.dtmainrekvizitler.Rows[0][5].ToString();
                    txtbankvoen.Text = MyData.dtmainrekvizitler.Rows[0][6].ToString();
                    txtmuxhesab1.Text = MyData.dtmainrekvizitler.Rows[0][7].ToString().Substring(0,8);
                    txtmuxhesab2.Text = MyData.dtmainrekvizitler.Rows[0][7].ToString().Substring(9, MyData.dtmainrekvizitler.Rows[0][7].ToString().Length-9);
                    txtsvift.Text = MyData.dtmainrekvizitler.Rows[0][8].ToString();
                }
                catch { MessageBox.Show("Tapılmadı"); };

                button22.Visible = true;
            }
        }

        private void txtmusteriadi_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1 Like '%" + txtmusteriadi.Text + "%'");
                MyData.dtmainrekvizitler=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

                try
                {
                    txtmusteriadi.Text = MyData.dtmainrekvizitler.Rows[0][1].ToString();
                    txthesab1.Text = MyData.dtmainrekvizitler.Rows[0][2].ToString().Substring(0,8);
                    txthesab2.Text = MyData.dtmainrekvizitler.Rows[0][2].ToString().Substring(9, MyData.dtmainrekvizitler.Rows[0][2].ToString().Length-9); ;
                    txtmusterivoen.Text = MyData.dtmainrekvizitler.Rows[0][3].ToString();
                    txtsexsiyyet.Text = MyData.dtmainrekvizitler.Rows[0][9].ToString();

                    txtbankadi.Text = MyData.dtmainrekvizitler.Rows[0][4].ToString();
                    txtbankkodu.Text = MyData.dtmainrekvizitler.Rows[0][5].ToString();
                    txtbankvoen.Text = MyData.dtmainrekvizitler.Rows[0][6].ToString();
                    txtmuxhesab1.Text = MyData.dtmainrekvizitler.Rows[0][7].ToString().Substring(0, 8);
                    txtmuxhesab2.Text = MyData.dtmainrekvizitler.Rows[0][7].ToString().Substring(9, MyData.dtmainrekvizitler.Rows[0][2].ToString().Length - 9); ;
                    txtsvift.Text = MyData.dtmainrekvizitler.Rows[0][8].ToString();
                }
                catch { MessageBox.Show("Tapılmadı"); };

                MyData.selectCommand("baza.accdb", "Select * From plateshkateyinat WHERE a1 Like " + "'%" + txtmusteriadi.Text + "%'");
                MyData.dtmainTeyinat=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainTeyinat);

                try
                {
                    txtesas.Text = MyData.dtmainTeyinat.Rows[0][1].ToString();
                }
                catch { };

                button22.Visible = true;
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            if (!MyCheck.davamYesNo()) return;

            try
            {
                if (checkBox1.Checked == true && txtmanatEmekhaqqi.Text.Substring(txtmanatEmekhaqqi.Text.Length - 3, 1) != ".")
                {
                    MessageBox.Show("Əmək haqqı hissəsi- Məbləğ düzgün qeyd olunmayıb." + Environment.NewLine + "Nümunə: 123.45"); return;
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); return; }

            try
            {
                if (checkBox2.Checked == true && txtmanat3faiz.Text.Substring(txtmanat3faiz.Text.Length - 3, 1) != ".")
                {
                    MessageBox.Show("10% - Məbləğ düzgün qeyd olunmayıb." + Environment.NewLine + "Nümunə: 123.45"); return;
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); return; }

            try
            {
                if (checkBox3.Checked == true && txtmanat22faiz.Text.Substring(txtmanat22faiz.Text.Length - 3, 1) != ".")
                {
                    MessageBox.Show("15% - Məbləğ düzgün qeyd olunmayıb." + Environment.NewLine + "Nümunə: 123.45"); return;
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); return; }

            try
            {
                if (checkBox4.Checked == true && txtmanatgelirvergisi.Text.Substring(txtmanatgelirvergisi.Text.Length - 3, 1) != ".")
                {
                    MessageBox.Show("Gəlir vergisi - Məbləğ düzgün qeyd olunmayıb." + Environment.NewLine + "Nümunə: 123.45"); return;
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); return; }

            try
            {
                if (checkBox5.Checked == true && txtmanatSigEden05faiz.Text.Substring(txtmanatSigEden05faiz.Text.Length - 3, 1) != ".")
                {
                    MessageBox.Show("Siğorta edən 0.5% - Məbləğ düzgün qeyd olunmayıb." + Environment.NewLine + "Nümunə: 123.45"); return;
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); return; }

            try
            {
                if (checkBox6.Checked == true && txtmanatSigolunan05faiz.Text.Substring(txtmanatSigolunan05faiz.Text.Length - 3, 1) != ".")
                {
                    MessageBox.Show("Siğorta olunan 0.5% - Məbləğ düzgün qeyd olunmayıb." + Environment.NewLine + "Nümunə: 123.45"); return;
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); return; }

            DateTime dt = dttarixEmekhaqqi.Value.Date;
            string a = MyChange.TarixSozle(dt);

            //EMEK HAQQİ--------------------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------------------------------------------
            if (checkBox1.Checked == true)
            {
                try { File.Copy("Plateshka.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Emek haqqi.xlsx", true); }
                catch { MessageBox.Show("'Plateshka.xlsx' tapılmadı."); }

                oXL = new Excel.Application();
                oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Emek haqqi.xlsx"));
                oSheet = (Excel._Worksheet)oWB.Sheets[1];
                oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                oXL.Visible = true;
                oSheet.Activate();
                oSheet.Range["A1"].Select();
                oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

                if (txtmanatEmekhaqqi2.Text == "") txtmanatEmekhaqqi2.Text = "00";
                if (txtmanatEmekhaqqi.Text == "") txtmanatEmekhaqqi.Text = "0";
                if (txtmanatEmekhaqqi2.Text == "") txtmanatEmekhaqqi2.Text = "00";
                if (txtmanat3faiz.Text == "") txtmanat3faiz.Text = "0";
                if (txtmanatEmekhaqqi6.Text == "") txtmanatEmekhaqqi6.Text = "00";
                if (txtmanat22faiz.Text == "") txtmanat22faiz.Text = "0";

                oSheet.Cells[2, 2] = "ÖDƏNİŞ TAPŞIRIĞI № " + txtnomreEmekhaqqi.Text + "/" + txtnomreEmekhaqqi2.Text;
                oSheet.Cells[4, 2] = dt.Day + " " + a + " " + dt.Year + " - ci il";
                oSheet.Cells[22, 8] = "AZN";
                oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + txtmanatEmekhaqqi.Text;
                if (txtmanatEmekhaqqi2.Text.Length < 2) oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + txtmanatEmekhaqqi.Text;

                oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: " + txtmanatEmekhaqqi2.Text;
                oSheet.Cells[28, 2] = "D1. Ödənişin təyinatı və əsas / Payment details: Əmək haqqı";
                oSheet.Cells[31, 2] = "D2. Ödənişlə əlaqədar əlavə informasiya / Narrative:";
                oSheet.Cells[35, 2] = "D3. Büdcə təsnifatının kodu:";
                oSheet.Cells[35, 8] = "D4. Büdcə səviyyəsinin kodu:";

                MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1='" + cbFromEmekHaqqi.Text + "'");
                MyData.dtmainrekvizitler=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

                string adi = MyData.dtmainrekvizitler.Rows[0][1].ToString();
                try
                {
                    for (int t = 0; t < MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; t++)
                    {
                        if (MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(t, 1) == "(") { adi = MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(0, t); t = MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; }
                    }
                }
                catch { }

                oSheet.Cells[17, 2] = "Adı / Name: " + adi;
                oSheet.Cells[18, 2] = "Hesab № / Acc. №: " + MyData.dtmainrekvizitler.Rows[0]["a2"].ToString();
                oSheet.Cells[20, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a3"].ToString();

                oSheet.Cells[8, 2] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0]["a4"].ToString();
                oSheet.Cells[9, 2] = "Kodu / Code: " + MyData.dtmainrekvizitler.Rows[0]["a5"].ToString();
                oSheet.Cells[10, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a6"].ToString();
                oSheet.Cells[11, 2] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainrekvizitler.Rows[0]["a7"].ToString();
                oSheet.Cells[13, 2] = "S. W. I. F. T. Bik: " + MyData.dtmainrekvizitler.Rows[0]["a8"].ToString();

                oSheet.Cells[17, 8] = "Adı / Name: “AGLizinq” QSC";
                oSheet.Cells[18, 8] = "Hesab № / Acc. №: AZ30AZEG" + Environment.NewLine + "45013944017950107055";
                oSheet.Cells[20, 8] = "VÖEN / Tax İD: 1300616961";

                oSheet.Cells[8, 8] = "Adı / Name: AGBANK ASC";
                oSheet.Cells[9, 8] = "Kodu / Code: 505817";
                oSheet.Cells[10, 8] = "VÖEN / Tax İD: 9900019651";
                oSheet.Cells[11, 8] = "Müxbir hesab / Corr. Acc.: AZ75NABZ" + Environment.NewLine + "01350100000000017944"; ;
                oSheet.Cells[13, 8] = "S. W. I. F. T. Bik: AZEGAZ22";

                MyData.updateCommand("baza.accdb", "UPDATE plateshkanomre SET nomre=" + "'" + txtnomreEmekhaqqi.Text + "'");

                MyData.insertCommand("baza.accdb", "insert into plateshkaARXIV (TARİX, №, HARDAN, HARA, [ÖDƏNİŞİN TƏYİNATI], MƏBLƏĞ, VALYUTA) Values ('" + dttarixEmekhaqqi.Text + "', '" + txtnomreEmekhaqqi.Text + "/" + txtnomreEmekhaqqi2.Text + "', '" + cbFromEmekHaqqi.Text + "', 'AGLizinq QSC', 'Əmək haqqı" + "', '" + txtmanatEmekhaqqi.Text + "', '" + "manat" + "')");

                MyData.insertCommand("baza.accdb", "insert into Emeliyyatlar (a1, a2, a3) Values ('" + DateTime.Now + "', 'ÖDƏNİŞ TAPŞIRIĞI - " + "ƏMƏK HAQQI" + " - " + txtmanatEmekhaqqi.Text + " - Əmək haqqı" + "','" + Environment.MachineName + "')");
                
                plateshkanomreRefresh();
            }

            //10%--------------------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------------------------------------------
            if (checkBox2.Checked == true)
            {
                try { File.Copy("Plateshka.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\10%.xlsx", true); }
                catch { MessageBox.Show("'Plateshka.xlsx' tapılmadı."); }

                oXL = new Excel.Application();
                oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\10%.xlsx"));
                oSheet = (Excel._Worksheet)oWB.Sheets[1];
                oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                oXL.Visible = true;
                oSheet.Activate();
                oSheet.Range["A1"].Select();
                oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

                if (txtmanatEmekhaqqi2.Text == "") txtmanatEmekhaqqi2.Text = "00";
                if (txtmanatEmekhaqqi.Text == "") txtmanatEmekhaqqi.Text = "0";
                if (txtmanat3faiz.Text == "") txtmanat3faiz.Text = "0";
                if (txtmanatEmekhaqqi6.Text == "") txtmanatEmekhaqqi6.Text = "00";
                if (txtmanat22faiz.Text == "") txtmanat22faiz.Text = "0";

                oSheet.Cells[2, 2] = "ÖDƏNİŞ TAPŞIRIĞI № " + txtnomreEmekhaqqi.Text + "/" + txtnomreEmekhaqqi2.Text;
                oSheet.Cells[4, 2] = dt.Day + " " + a + " " + dt.Year + " - ci il";
                oSheet.Cells[22, 8] = "AZN";
                oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + txtmanat3faiz.Text;
                if (txtmanatEmekhaqqi2.Text.Length < 2) oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + txtmanat3faiz.Text + ".0" + txtmanatEmekhaqqi2.Text;

                oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: " + txtmanatEmekhaqqi4.Text;
                oSheet.Cells[28, 2] = "D1. Ödənişin təyinatı və əsas / Payment details: DSMF-na ayırmalar - (10%)";
                oSheet.Cells[31, 2] = "D2. Ödənişlə əlaqədar əlavə informasiya / Narrative:";
                oSheet.Cells[35, 2] = "D3. Büdcə təsnifatının kodu: 121211";
                oSheet.Cells[35, 8] = "D4. Büdcə səviyyəsinin kodu: 4";

                MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1='" + cbFromEmekHaqqi.Text + "'");
                MyData.dtmainrekvizitler = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

                string adi = MyData.dtmainrekvizitler.Rows[0][1].ToString();
                try
                {
                    for (int t = 0; t < MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; t++)
                    {
                        if (MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(t, 1) == "(") { adi = MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(0, t); t = MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; }
                    }
                }
                catch { }

                oSheet.Cells[17, 2] = "Adı / Name: " + adi;
                oSheet.Cells[18, 2] = "Hesab № / Acc. №: " + MyData.dtmainrekvizitler.Rows[0]["a2"].ToString();
                oSheet.Cells[20, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a3"].ToString();

                oSheet.Cells[8, 2] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0]["a4"].ToString();
                oSheet.Cells[9, 2] = "Kodu / Code: " + MyData.dtmainrekvizitler.Rows[0]["a5"].ToString();
                oSheet.Cells[10, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a6"].ToString();
                oSheet.Cells[11, 2] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainrekvizitler.Rows[0]["a7"].ToString();
                oSheet.Cells[13, 2] = "S. W. I. F. T. Bik: " + MyData.dtmainrekvizitler.Rows[0]["a8"].ToString();
      
                oSheet.Cells[17, 8] = "Adı / Name: DSMF-nin Mərkəzi Sosial Siğorta və Fərdi Uçot İdarəsi";
                oSheet.Cells[18, 8] = "Hesab № / Acc. №: AZ89CTRE" + Environment.NewLine + "00000000000007018506";
                oSheet.Cells[20, 8] = "VÖEN / Tax İD: 1300115511";

                oSheet.Cells[8, 8] = "Adı / Name: DXA";
                oSheet.Cells[9, 8] = "Kodu / Code: 210005";
                oSheet.Cells[10, 8] = "VÖEN / Tax İD: 1401555071";
                oSheet.Cells[11, 8] = "Müxbir hesab / Corr. Acc.: AZ41NABZ" + Environment.NewLine + "01360100000000003944"; ;
                oSheet.Cells[13, 8] = "S. W. I. F. T. Bik: CTREAZ22";

                MyData.updateCommand("baza.accdb", "UPDATE plateshkanomre SET nomre=" + "'" + txtnomreEmekhaqqi.Text + "'");
                
                MyData.insertCommand("baza.accdb", "insert into plateshkaARXIV (TARİX, №, HARDAN, HARA, [ÖDƏNİŞİN TƏYİNATI], MƏBLƏĞ, VALYUTA) Values ('" + dttarixEmekhaqqi.Text + "', '" + txtnomreEmekhaqqi.Text + "/" + txtnomreEmekhaqqi2.Text + "', '" + cbFromEmekHaqqi.Text + "', 'DSMF-nin Mərkəzi Sosial Siğorta və Fərdi Uçot İdarəsi', 'DSMF-na ayırmalar (10%)" + "', '" + txtmanat3faiz.Text + "', '" + "manat" + "')");
                
                MyData.insertCommand("baza.accdb", "insert into Emeliyyatlar (a1, a2, a3) Values ('" + DateTime.Now + "', 'ÖDƏNİŞ TAPŞIRIĞI - " + "ƏMƏK HAQQI" + " - " + txtmanat3faiz.Text + " - DSMF 10%" + "','" + Environment.MachineName + "')");
                
                plateshkanomreRefresh();
            }

            //15%--------------------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------------------------------------------
            if (checkBox3.Checked == true)
            {
                try { File.Copy("Plateshka.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\15%.xlsx", true); }
                catch { MessageBox.Show("'Plateshka.xlsx' tapılmadı."); }

                oXL = new Excel.Application();
                oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\15%.xlsx"));
                oSheet = (Excel._Worksheet)oWB.Sheets[1];
                oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                oXL.Visible = true;
                oSheet.Activate();
                oSheet.Range["A1"].Select();
                oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

                if (txtmanatEmekhaqqi2.Text == "") txtmanatEmekhaqqi2.Text = "00";
                if (txtmanatEmekhaqqi.Text == "") txtmanatEmekhaqqi.Text = "0";
                if (txtmanat3faiz.Text == "") txtmanat3faiz.Text = "0";
                if (txtmanatEmekhaqqi6.Text == "") txtmanatEmekhaqqi6.Text = "00";
                if (txtmanat22faiz.Text == "") txtmanat22faiz.Text = "0";

                oSheet.Cells[2, 2] = "ÖDƏNİŞ TAPŞIRIĞI № " + txtnomreEmekhaqqi.Text + "/" + txtnomreEmekhaqqi2.Text;
                oSheet.Cells[4, 2] = dt.Day + " " + a + " " + dt.Year + " - ci il";
                oSheet.Cells[22, 8] = "AZN";
                oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + txtmanat22faiz.Text;
                if (txtmanat22faiz.Text.Length < 2) oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + txtmanat22faiz.Text + ".0" + txtmanatEmekhaqqi6.Text;

                oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: " + txtmanatEmekhaqqi6.Text;
                oSheet.Cells[28, 2] = "D1. Ödənişin təyinatı və əsas / Payment details: DSMF-na ayırmalar - (15%)";
                oSheet.Cells[31, 2] = "D2. Ödənişlə əlaqədar əlavə informasiya / Narrative:";
                oSheet.Cells[35, 2] = "D3. Büdcə təsnifatının kodu: 121111";
                oSheet.Cells[35, 8] = "D4. Büdcə səviyyəsinin kodu: 4";

                MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1='" + cbFromEmekHaqqi.Text + "'");
                MyData.dtmainrekvizitler=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

                string adi = MyData.dtmainrekvizitler.Rows[0][1].ToString();
                try
                {
                    for (int t = 0; t < MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; t++)
                    {
                        if (MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(t, 1) == "(") { adi = MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(0, t); t = MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; }
                    }
                }
                catch { }

                oSheet.Cells[17, 2] = "Adı / Name: " + adi;
                oSheet.Cells[18, 2] = "Hesab № / Acc. №: " + MyData.dtmainrekvizitler.Rows[0]["a2"].ToString();
                oSheet.Cells[20, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a3"].ToString();

                oSheet.Cells[8, 2] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0]["a4"].ToString();
                oSheet.Cells[9, 2] = "Kodu / Code: " + MyData.dtmainrekvizitler.Rows[0]["a5"].ToString();
                oSheet.Cells[10, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a6"].ToString();
                oSheet.Cells[11, 2] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainrekvizitler.Rows[0]["a7"].ToString();
                oSheet.Cells[13, 2] = "S. W. I. F. T. Bik: " + MyData.dtmainrekvizitler.Rows[0]["a8"].ToString();

                oSheet.Cells[17, 8] = "Adı / Name: DSMF-nin Mərkəzi Sosial Siğorta və Fərdi Uçot İdarəsi";
                oSheet.Cells[18, 8] = "Hesab № / Acc. №: AZ89CTRE" + Environment.NewLine + "00000000000007018506";
                oSheet.Cells[20, 8] = "VÖEN / Tax İD: 1300115511";

                oSheet.Cells[8, 8] = "Adı / Name: DXA";
                oSheet.Cells[9, 8] = "Kodu / Code: 210005";
                oSheet.Cells[10, 8] = "VÖEN / Tax İD: 1401555071";
                oSheet.Cells[11, 8] = "Müxbir hesab / Corr. Acc.: AZ41NABZ" + Environment.NewLine + "01360100000000003944"; ;
                oSheet.Cells[13, 8] = "S. W. I. F. T. Bik: CTREAZ22";

                MyData.updateCommand("baza.accdb", "UPDATE plateshkanomre SET nomre=" + "'" + txtnomreEmekhaqqi.Text + "'");
                
                MyData.insertCommand("baza.accdb", "insert into plateshkaARXIV (TARİX, №, HARDAN, HARA, [ÖDƏNİŞİN TƏYİNATI], MƏBLƏĞ, VALYUTA) Values ('" + dttarixEmekhaqqi.Text + "', '" + txtnomreEmekhaqqi.Text + "/" + txtnomreEmekhaqqi2.Text + "', '" + cbFromEmekHaqqi.Text + "', 'DSMF-nin Mərkəzi Sosial Siğorta və Fərdi Uçot İdarəsi', 'DSMF-na ayırmalar (15%)" + "', '" + txtmanat22faiz.Text + "', '" + "manat" + "')");
                
                MyData.insertCommand("baza.accdb", "insert into Emeliyyatlar (a1, a2, a3) Values ('" + DateTime.Now + "', 'ÖDƏNİŞ TAPŞIRIĞI - " + "ƏMƏK HAQQI" + " - " + txtmanat22faiz.Text + " - DSMF 15%" + "','" + Environment.MachineName + "')");
                
                plateshkanomreRefresh();
            }

            //---Sigorta Eden ucun 05%-------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------------------------------------------

            if (checkBox5.Checked == true)
            {
                try { File.Copy("Plateshka.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Sigorta Eden 05%.xlsx", true); }
                catch { MessageBox.Show("'Plateshka.xlsx' tapılmadı."); }

                oXL = new Excel.Application();
                oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Sigorta Eden 05%.xlsx"));
                oSheet = (Excel._Worksheet)oWB.Sheets[1];
                oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                oXL.Visible = true;
                oSheet.Activate();
                oSheet.Range["A1"].Select();
                oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

                if (txtmanatEmekhaqqi.Text == "") txtmanatEmekhaqqi.Text = "0.00";
                if (txtmanat3faiz.Text == "") txtmanat3faiz.Text = "0.00";
                if (txtmanat22faiz.Text == "") txtmanat22faiz.Text = "0.00";
                if (txtmanatgelirvergisi.Text == "") txtmanatgelirvergisi.Text = "0.00";
                if (txtmanatSigEden05faiz.Text == "") txtmanatSigEden05faiz.Text = "0.00";
                if (txtmanatSigolunan05faiz.Text == "") txtmanatSigolunan05faiz.Text = "0.00";

                oSheet.Cells[2, 2] = "ÖDƏNİŞ TAPŞIRIĞI № " + txtnomreEmekhaqqi.Text + "/" + txtnomreEmekhaqqi2.Text;
                oSheet.Cells[4, 2] = dt.Day + " " + a + " " + dt.Year + " - ci il";
                oSheet.Cells[22, 8] = "AZN";
                oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + txtmanatSigEden05faiz.Text;
                oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: " + txtmanatEmekhaqqi10.Text;
                oSheet.Cells[28, 2] = "D1. Ödənişin təyinatı və əsas / Payment details: İşəgötürənlərin işsizlikdən siğorta haqqı - (0.5%)";
                oSheet.Cells[31, 2] = "D2. Ödənişlə əlaqədar əlavə informasiya / Narrative:";
                oSheet.Cells[35, 2] = "D3. Büdcə təsnifatının kodu: 123100";
                oSheet.Cells[35, 8] = "D4. Büdcə səviyyəsinin kodu: 4";

                MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1='" + cbFromEmekHaqqi.Text + "'");
                MyData.dtmainrekvizitler=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

                string adi = MyData.dtmainrekvizitler.Rows[0][1].ToString();
                try
                {
                    for (int t = 0; t < MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; t++)
                    {
                        if (MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(t, 1) == "(") { adi = MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(0, t); t = MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; }
                    }
                }
                catch { }

                oSheet.Cells[17, 2] = "Adı / Name: " + adi;
                oSheet.Cells[18, 2] = "Hesab № / Acc. №: " + MyData.dtmainrekvizitler.Rows[0]["a2"].ToString();
                oSheet.Cells[20, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a3"].ToString();

                oSheet.Cells[8, 2] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0]["a4"].ToString();
                oSheet.Cells[9, 2] = "Kodu / Code: " + MyData.dtmainrekvizitler.Rows[0]["a5"].ToString();
                oSheet.Cells[10, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a6"].ToString();
                oSheet.Cells[11, 2] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainrekvizitler.Rows[0]["a7"].ToString();
                oSheet.Cells[13, 2] = "S. W. I. F. T. Bik: " + MyData.dtmainrekvizitler.Rows[0]["a8"].ToString();

                oSheet.Cells[17, 8] = "Adı / Name: DSMF-nin Mərkəzi Sosial Siğorta və Fərdi Uçot İdarəsi";
                oSheet.Cells[18, 8] = "Hesab № / Acc. №: AZ53CTRE" + Environment.NewLine + "00000000000007018572";
                oSheet.Cells[20, 8] = "VÖEN / Tax İD: 1300115511";

                oSheet.Cells[8, 8] = "Adı / Name: DXA";
                oSheet.Cells[9, 8] = "Kodu / Code: 210005";
                oSheet.Cells[10, 8] = "VÖEN / Tax İD: 1401555071";
                oSheet.Cells[11, 8] = "Müxbir hesab / Corr. Acc.: AZ41NABZ" + Environment.NewLine + "01360100000000003944"; ;
                oSheet.Cells[13, 8] = "S. W. I. F. T. Bik: CTREAZ22";

                MyData.updateCommand("baza.accdb", "UPDATE plateshkanomre SET nomre=" + "'" + txtnomreEmekhaqqi.Text + "'");
                
                MyData.insertCommand("baza.accdb", "insert into plateshkaARXIV (TARİX, №, HARDAN, HARA, [ÖDƏNİŞİN TƏYİNATI], MƏBLƏĞ, VALYUTA) Values ('" + dttarixEmekhaqqi.Text + "', '" + txtnomreEmekhaqqi.Text + "/" + txtnomreEmekhaqqi2.Text + "', '" + cbFromEmekHaqqi.Text + "', 'DSMF-nin Mərkəzi Sosial Siğorta və Fərdi Uçot İdarəsi', 'İşəgötürənlərin işsizlikdən siğorta haqqı - (0.5%)" + "', '" + txtmanatSigEden05faiz.Text + "', '" + "manat" + "')");
                
                MyData.insertCommand("baza.accdb", "insert into Emeliyyatlar (a1, a2, a3) Values ('" + DateTime.Now + "', 'ÖDƏNİŞ TAPŞIRIĞI - " + "ƏMƏK HAQQI" + " - " + txtmanatSigEden05faiz.Text + " - İşəgötürənlərin işsizlikdən siğorta haqqı - (0.5%)" + "','" + Environment.MachineName + "')");
                
                plateshkanomreRefresh();
            }

            //---Sigorta Olunan ucun 05%-------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------------------------------------------

            if (checkBox6.Checked == true)
            {
                try { File.Copy("Plateshka.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Sigorta Olunan 05%.xlsx", true); }
                catch { MessageBox.Show("'Plateshka.xlsx' tapılmadı."); }

                oXL = new Excel.Application();
                oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Sigorta Olunan 05%.xlsx"));
                oSheet = (Excel._Worksheet)oWB.Sheets[1];
                oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                oXL.Visible = true;
                oSheet.Activate();
                oSheet.Range["A1"].Select();
                oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

                if (txtmanatEmekhaqqi.Text == "") txtmanatEmekhaqqi.Text = "0.00";
                if (txtmanat3faiz.Text == "") txtmanat3faiz.Text = "0.00";
                if (txtmanat22faiz.Text == "") txtmanat22faiz.Text = "0.00";
                if (txtmanatgelirvergisi.Text == "") txtmanatgelirvergisi.Text = "0.00";
                if (txtmanatSigEden05faiz.Text == "") txtmanatSigEden05faiz.Text = "0.00";
                if (txtmanatSigolunan05faiz.Text == "") txtmanatSigolunan05faiz.Text = "0.00";

                oSheet.Cells[2, 2] = "ÖDƏNİŞ TAPŞIRIĞI № " + txtnomreEmekhaqqi.Text + "/" + txtnomreEmekhaqqi2.Text;
                oSheet.Cells[4, 2] = dt.Day + " " + a + " " + dt.Year + " - ci il";
                oSheet.Cells[22, 8] = "AZN";
                oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + txtmanatSigolunan05faiz.Text;
                oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: " + txtmanatEmekhaqqi12.Text;
                oSheet.Cells[28, 2] = "D1. Ödənişin təyinatı və əsas / Payment details: İşləyənlərin işsizlikdən siğorta haqqı - (0.5%)";
                oSheet.Cells[31, 2] = "D2. Ödənişlə əlaqədar əlavə informasiya / Narrative:";
                oSheet.Cells[35, 2] = "D3. Büdcə təsnifatının kodu: 123200";
                oSheet.Cells[35, 8] = "D4. Büdcə səviyyəsinin kodu: 4";

                MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1='" + cbFromEmekHaqqi.Text + "'");
                MyData.dtmainrekvizitler=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

                string adi = MyData.dtmainrekvizitler.Rows[0][1].ToString();
                try
                {
                    for (int t = 0; t < MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; t++)
                    {
                        if (MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(t, 1) == "(") { adi = MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(0, t); t = MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; }
                    }
                }
                catch { }

                oSheet.Cells[17, 2] = "Adı / Name: " + adi;
                oSheet.Cells[18, 2] = "Hesab № / Acc. №: " + MyData.dtmainrekvizitler.Rows[0]["a2"].ToString();
                oSheet.Cells[20, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a3"].ToString();

                oSheet.Cells[8, 2] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0]["a4"].ToString();
                oSheet.Cells[9, 2] = "Kodu / Code: " + MyData.dtmainrekvizitler.Rows[0]["a5"].ToString();
                oSheet.Cells[10, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a6"].ToString();
                oSheet.Cells[11, 2] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainrekvizitler.Rows[0]["a7"].ToString();
                oSheet.Cells[13, 2] = "S. W. I. F. T. Bik: " + MyData.dtmainrekvizitler.Rows[0]["a8"].ToString();

                oSheet.Cells[17, 8] = "Adı / Name: DSMF-nin Mərkəzi Sosial Siğorta və Fərdi Uçot İdarəsi";
                oSheet.Cells[18, 8] = "Hesab № / Acc. №: AZ53CTRE" + Environment.NewLine + "00000000000007018572";
                oSheet.Cells[20, 8] = "VÖEN / Tax İD: 1300115511";

                oSheet.Cells[8, 8] = "Adı / Name: DXA";
                oSheet.Cells[9, 8] = "Kodu / Code: 210005";
                oSheet.Cells[10, 8] = "VÖEN / Tax İD: 1401555071";
                oSheet.Cells[11, 8] = "Müxbir hesab / Corr. Acc.: AZ41NABZ" + Environment.NewLine + "01360100000000003944"; ;
                oSheet.Cells[13, 8] = "S. W. I. F. T. Bik: CTREAZ22";

                MyData.updateCommand("baza.accdb", "UPDATE plateshkanomre SET nomre=" + "'" + txtnomreEmekhaqqi.Text + "'");
                
                MyData.insertCommand("baza.accdb", "insert into plateshkaARXIV (TARİX, №, HARDAN, HARA, [ÖDƏNİŞİN TƏYİNATI], MƏBLƏĞ, VALYUTA) Values ('" + dttarixEmekhaqqi.Text + "', '" + txtnomreEmekhaqqi.Text + "/" + txtnomreEmekhaqqi2.Text + "', '" + cbFromEmekHaqqi.Text + "', 'DSMF-nin Mərkəzi Sosial Siğorta və Fərdi Uçot İdarəsi', 'İşləyənlərin işsizlikdən siğorta haqqı - (0.5%)" + "', '" + txtmanatSigolunan05faiz.Text + "', '" + "manat" + "')");
                
                MyData.insertCommand("baza.accdb", "insert into Emeliyyatlar (a1, a2, a3) Values ('" + DateTime.Now + "', 'ÖDƏNİŞ TAPŞIRIĞI - " + "ƏMƏK HAQQI" + " - " + txtmanatSigolunan05faiz.Text + " - İşləyənlərin işsizlikdən siğorta haqqı - (0.5%)" + "','" + Environment.MachineName + "')");
                
                plateshkanomreRefresh();
            }

            //---Gelir vergisi-------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------------------------------------------
                if (checkBox4.Checked == true)
                {
                    try { File.Copy("Plateshka.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\14%.xlsx", true); }
                    catch { MessageBox.Show("'Plateshka.xlsx' tapılmadı."); }

                    oXL = new Excel.Application();
                    oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\14%.xlsx"));
                    oSheet = (Excel._Worksheet)oWB.Sheets[1];
                    oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                    oXL.Visible = true;
                    oSheet.Activate();
                    oSheet.Range["A1"].Select();
                    oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

                    if (txtmanatEmekhaqqi2.Text == "") txtmanatEmekhaqqi2.Text = "00";
                    if (txtmanatEmekhaqqi.Text == "") txtmanatEmekhaqqi.Text = "0";
                    if (txtmanatEmekhaqqi2.Text == "") txtmanatEmekhaqqi2.Text = "00";
                    if (txtmanat3faiz.Text == "") txtmanat3faiz.Text = "0";
                    if (txtmanatEmekhaqqi6.Text == "") txtmanatEmekhaqqi6.Text = "00";
                    if (txtmanat22faiz.Text == "") txtmanat22faiz.Text = "0";
                    if (txtmanatEmekhaqqi8.Text == "") txtmanatEmekhaqqi6.Text = "00";
                    if (txtmanatgelirvergisi.Text == "") txtmanat22faiz.Text = "0";

                    oSheet.Cells[2, 2] = "ÖDƏNİŞ TAPŞIRIĞI № " + txtnomreEmekhaqqi.Text + "/" + txtnomreEmekhaqqi2.Text;
                    oSheet.Cells[4, 2] = dt.Day + " " + a + " " + dt.Year + " - ci il";
                    oSheet.Cells[22, 8] = "AZN";
                    oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + txtmanatgelirvergisi.Text;
                    oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: " + txtmanatEmekhaqqi8.Text;
                    oSheet.Cells[28, 2] = "D1. Ödənişin təyinatı və əsas / Payment details: " + Environment.NewLine + textBox8.Text;
                    oSheet.Cells[31, 2] = "D2. Ödənişlə əlaqədar əlavə informasiya / Narrative:";
                    oSheet.Cells[35, 2] = "D3. Büdcə təsnifatının kodu: 111111";
                    oSheet.Cells[35, 8] = "D4. Büdcə səviyyəsinin kodu: 1";

                    MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1='" + cbFromEmekHaqqi.Text + "'");
                    MyData.dtmainrekvizitler=new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

                    string adi = MyData.dtmainrekvizitler.Rows[0][1].ToString();
                    try
                    {
                        for (int t = 0; t < MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; t++)
                        {
                            if (MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(t, 1) == "(") { adi = MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(0, t); t = MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; }
                        }
                    }
                    catch { }

                    oSheet.Cells[17, 2] = "Adı / Name: " + adi;
                    oSheet.Cells[18, 2] = "Hesab № / Acc. №: " + MyData.dtmainrekvizitler.Rows[0]["a2"].ToString();
                    oSheet.Cells[20, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a3"].ToString();

                    oSheet.Cells[8, 2] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0]["a4"].ToString();
                    oSheet.Cells[9, 2] = "Kodu / Code: " + MyData.dtmainrekvizitler.Rows[0]["a5"].ToString();
                    oSheet.Cells[10, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a6"].ToString();
                    oSheet.Cells[11, 2] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainrekvizitler.Rows[0]["a7"].ToString();
                    oSheet.Cells[13, 2] = "S. W. I. F. T. Bik: " + MyData.dtmainrekvizitler.Rows[0]["a8"].ToString();

                    oSheet.Cells[17, 8] = "Adı / Name: Bakı şəhəri Lokal gəlirlər Departamenti";
                    oSheet.Cells[18, 8] = "Hesab № / Acc. №: AZ17CTRE" + Environment.NewLine + "00000000000002117131";
                    oSheet.Cells[20, 8] = "VÖEN / Tax İD: 1403006271";

                    oSheet.Cells[8, 8] = "Adı / Name: DXA";
                    oSheet.Cells[9, 8] = "Kodu / Code: 210005";
                    oSheet.Cells[10, 8] = "VÖEN / Tax İD: 1401555071";
                    oSheet.Cells[11, 8] = "Müxbir hesab / Corr. Acc.: AZ41NABZ" + Environment.NewLine + "01360100000000003944"; ;
                    oSheet.Cells[13, 8] = "S. W. I. F. T. Bik: CTREAZ22";

                    MyData.updateCommand("baza.accdb", "UPDATE plateshkanomre SET nomre=" + "'" + txtnomreEmekhaqqi.Text + "'");

                    MyData.insertCommand("baza.accdb", "insert into plateshkaARXIV (TARİX, №, HARDAN, HARA, [ÖDƏNİŞİN TƏYİNATI], MƏBLƏĞ, VALYUTA) Values ('" + dttarixEmekhaqqi.Text + "', '" + txtnomreEmekhaqqi.Text + "/" + txtnomreEmekhaqqi2.Text + "', '“AGLizinq” QSC (AGB AZN)', 'Bakı şəhəri Lokal gəlirlər Departamenti" + "', '" + textBox8.Text + "', '" + txtmanatgelirvergisi.Text + "', '" + "manat" + "')");
                    
                    MyData.insertCommand("baza.accdb", "insert into Emeliyyatlar (a1, a2, a3) Values ('" + DateTime.Now + "', 'ÖDƏNİŞ TAPŞIRIĞI - " + "Bakı şəhəri Lokal gəlirlər Departamenti" + " - " + txtmanatgelirvergisi.Text + " (" + textBox8.Text + ")','" + Environment.MachineName + "')");
                    
                    plateshkanomreRefresh();
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            if (!MyCheck.davamYesNo()) return;

            try
            {
                if (checkBox1.Checked == true && txtmanatEmekhaqqi.Text.Substring(txtmanatEmekhaqqi.Text.Length - 3, 1) != ".")
                {
                    MessageBox.Show("Əmək haqqı hissəsi- Məbləğ düzgün qeyd olunmayıb." + Environment.NewLine + "Nümunə: 123.45"); return;
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); return; }

            try
            {
                if (checkBox2.Checked == true && txtmanat3faiz.Text.Substring(txtmanat3faiz.Text.Length - 3, 1) != ".")
                {
                    MessageBox.Show("10% - Məbləğ düzgün qeyd olunmayıb." + Environment.NewLine + "Nümunə: 123.45"); return;
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); return; }

            try
            {
                if (checkBox3.Checked == true && txtmanat22faiz.Text.Substring(txtmanat22faiz.Text.Length - 3, 1) != ".")
                {
                    MessageBox.Show("15% - Məbləğ düzgün qeyd olunmayıb." + Environment.NewLine + "Nümunə: 123.45"); return;
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); return; }

            try
            {
                if (checkBox4.Checked == true && txtmanatgelirvergisi.Text.Substring(txtmanatgelirvergisi.Text.Length - 3, 1) != ".")
                {
                    MessageBox.Show("Gəlir vergisi - Məbləğ düzgün qeyd olunmayıb." + Environment.NewLine + "Nümunə: 123.45"); return;
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); return; }

            try
            {
                if (checkBox5.Checked == true && txtmanatSigEden05faiz.Text.Substring(txtmanatSigEden05faiz.Text.Length - 3, 1) != ".")
                {
                    MessageBox.Show("Siğorta edən 0.5% - Məbləğ düzgün qeyd olunmayıb." + Environment.NewLine + "Nümunə: 123.45"); return;
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); return; }

            try
            {
                if (checkBox6.Checked == true && txtmanatSigolunan05faiz.Text.Substring(txtmanatSigolunan05faiz.Text.Length - 3, 1) != ".")
                {
                    MessageBox.Show("Siğorta olunan 0.5% - Məbləğ düzgün qeyd olunmayıb." + Environment.NewLine + "Nümunə: 123.45"); return;
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); return; }

            DateTime dt = dttarixEmekhaqqi.Value.Date;
            string a = MyChange.TarixSozle(dt);
            
            //EMEK HAQQİ--------------------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------------------------------------------
            if (checkBox1.Checked == true)
            {
                try { File.Copy("Plateshka.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Emek haqqi.xlsx", true); }
                catch { MessageBox.Show("'Plateshka.xlsx' tapılmadı."); }

                oXL = new Excel.Application();
                oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Emek haqqi.xlsx"));
                oSheet = (Excel._Worksheet)oWB.Sheets[1];
                oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                oXL.Visible = false;
                oSheet.Activate();
                oSheet.Range["A1"].Select();
                oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

                if (txtmanatEmekhaqqi2.Text == "") txtmanatEmekhaqqi2.Text = "00";
                if (txtmanatEmekhaqqi.Text == "") txtmanatEmekhaqqi.Text = "0";
                if (txtmanatEmekhaqqi2.Text == "") txtmanatEmekhaqqi2.Text = "00";
                if (txtmanat3faiz.Text == "") txtmanat3faiz.Text = "0";
                if (txtmanatEmekhaqqi6.Text == "") txtmanatEmekhaqqi6.Text = "00";
                if (txtmanat22faiz.Text == "") txtmanat22faiz.Text = "0";

                oSheet.Cells[2, 2] = "ÖDƏNİŞ TAPŞIRIĞI № " + txtnomreEmekhaqqi.Text + "/" + txtnomreEmekhaqqi2.Text;
                oSheet.Cells[4, 2] = dt.Day + " " + a + " " + dt.Year + " - ci il";

                oSheet.Cells[22, 8] = "AZN";

                //Emek haqqi ucun
                oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + txtmanatEmekhaqqi.Text;
                if (txtmanatEmekhaqqi2.Text.Length < 2) oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + txtmanatEmekhaqqi.Text + ".0" + txtmanatEmekhaqqi2.Text;

                oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: " + txtmanatEmekhaqqi2.Text;
                oSheet.Cells[28, 2] = "D1. Ödənişin təyinatı və əsas / Payment details: Əmək haqqı";
                oSheet.Cells[31, 2] = "D2. Ödənişlə əlaqədar əlavə informasiya / Narrative:";
                oSheet.Cells[35, 2] = "D3. Büdcə təsnifatının kodu:";
                oSheet.Cells[35, 8] = "D4. Büdcə səviyyəsinin kodu:";

                MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1='" + cbFromEmekHaqqi.Text + "'");
                MyData.dtmainrekvizitler=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

                string adi = MyData.dtmainrekvizitler.Rows[0][1].ToString();
                try
                {
                    for (int t = 0; t < MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; t++)
                    {
                        if (MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(t, 1) == "(") { adi = MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(0, t); t = MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; }
                    }
                }
                catch { }

                oSheet.Cells[17, 2] = "Adı / Name: " + adi;
                oSheet.Cells[18, 2] = "Hesab № / Acc. №: " + MyData.dtmainrekvizitler.Rows[0]["a2"].ToString();
                oSheet.Cells[20, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a3"].ToString();

                oSheet.Cells[8, 2] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0]["a4"].ToString();
                oSheet.Cells[9, 2] = "Kodu / Code: " + MyData.dtmainrekvizitler.Rows[0]["a5"].ToString();
                oSheet.Cells[10, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a6"].ToString();
                oSheet.Cells[11, 2] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainrekvizitler.Rows[0]["a7"].ToString();
                oSheet.Cells[13, 2] = "S. W. I. F. T. Bik: " + MyData.dtmainrekvizitler.Rows[0]["a8"].ToString();

                oSheet.Cells[17, 8] = "Adı / Name: “AGLizinq” QSC";
                oSheet.Cells[18, 8] = "Hesab № / Acc. №: AZ30AZEG" + Environment.NewLine + "45013944017950107055";
                oSheet.Cells[20, 8] = "VÖEN / Tax İD: 1300616961";

                oSheet.Cells[8, 8] = "Adı / Name: AGBANK ASC";
                oSheet.Cells[9, 8] = "Kodu / Code: 505817";
                oSheet.Cells[10, 8] = "VÖEN / Tax İD: 9900019651";
                oSheet.Cells[11, 8] = "Müxbir hesab / Corr. Acc.: AZ75NABZ" + Environment.NewLine + "01350100000000017944"; ;
                oSheet.Cells[13, 8] = "S. W. I. F. T. Bik: AZEGAZ22";

                MyData.updateCommand("baza.accdb", "UPDATE plateshkanomre SET nomre=" + "'" + txtnomreEmekhaqqi.Text + "'");
                
                MyData.insertCommand("baza.accdb", "insert into plateshkaARXIV (TARİX, №, HARDAN, HARA, [ÖDƏNİŞİN TƏYİNATI], MƏBLƏĞ, VALYUTA) Values ('" + dttarixEmekhaqqi.Text + "', '" + txtnomreEmekhaqqi.Text + "/" + txtnomreEmekhaqqi2.Text + "', '" + cbFromEmekHaqqi.Text + "', 'AGLizinq QSC', 'Əmək haqqı" + "', '" + txtmanatEmekhaqqi.Text + "', '" + "manat" + "')");
                
                MyData.insertCommand("baza.accdb", "insert into Emeliyyatlar (a1, a2, a3) Values ('" + DateTime.Now + "', 'ÖDƏNİŞ TAPŞIRIĞI - " + "ƏMƏK HAQQI" + " - " + txtmanatEmekhaqqi.Text + " - Əmək haqqı" + "','" + Environment.MachineName + "')");
                
                oSheet.PrintOut(1, 1, 3);
                oXL.DisplayAlerts = false; 
                oWB.Close(SaveChanges: true);
                oXL.Application.Quit();
                plateshkanomreRefresh();
            }

            //10%--------------------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------------------------------------------
            if (checkBox2.Checked == true)
            {
                try { File.Copy("Plateshka.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\10%.xlsx", true); }
                catch { MessageBox.Show("'Plateshka.xlsx' tapılmadı."); }

                oXL = new Excel.Application();
                oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\10%.xlsx"));
                oSheet = (Excel._Worksheet)oWB.Sheets[1];
                oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                oXL.Visible = false;
                oSheet.Activate();
                oSheet.Range["A1"].Select();
                oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

                if (txtmanatEmekhaqqi2.Text == "") txtmanatEmekhaqqi2.Text = "00";
                if (txtmanatEmekhaqqi.Text == "") txtmanatEmekhaqqi.Text = "0";
                if (txtmanatEmekhaqqi2.Text == "") txtmanatEmekhaqqi2.Text = "00";
                if (txtmanat3faiz.Text == "") txtmanat3faiz.Text = "0";
                if (txtmanatEmekhaqqi6.Text == "") txtmanatEmekhaqqi6.Text = "00";
                if (txtmanat22faiz.Text == "") txtmanat22faiz.Text = "0";

                oSheet.Cells[2, 2] = "ÖDƏNİŞ TAPŞIRIĞI № " + txtnomreEmekhaqqi.Text + "/" + txtnomreEmekhaqqi2.Text;
                oSheet.Cells[4, 2] = dt.Day + " " + a + " " + dt.Year+ " - ci il";
                oSheet.Cells[22, 8] = "AZN";
                oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + txtmanat3faiz.Text;
                if (txtmanatEmekhaqqi2.Text.Length < 2) oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + txtmanat3faiz.Text + ".0" + txtmanatEmekhaqqi2.Text;

                oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: " + txtmanatEmekhaqqi4.Text;
                oSheet.Cells[28, 2] = "D1. Ödənişin təyinatı və əsas / Payment details: DSMF-na ayırmalar - (10%)";
                oSheet.Cells[31, 2] = "D2. Ödənişlə əlaqədar əlavə informasiya / Narrative:";
                oSheet.Cells[35, 2] = "D3. Büdcə təsnifatının kodu: 121211";
                oSheet.Cells[35, 8] = "D4. Büdcə səviyyəsinin kodu: 4";

                MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1='" + cbFromEmekHaqqi.Text + "'");
                MyData.dtmainrekvizitler=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

                string adi = MyData.dtmainrekvizitler.Rows[0][1].ToString();
                try
                {
                    for (int t = 0; t < MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; t++)
                    {
                        if (MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(t, 1) == "(") { adi = MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(0, t); t = MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; }
                    }
                }
                catch { }

                oSheet.Cells[17, 2] = "Adı / Name: " + adi;
                oSheet.Cells[18, 2] = "Hesab № / Acc. №: " + MyData.dtmainrekvizitler.Rows[0]["a2"].ToString();
                oSheet.Cells[20, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a3"].ToString();

                oSheet.Cells[8, 2] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0]["a4"].ToString();
                oSheet.Cells[9, 2] = "Kodu / Code: " + MyData.dtmainrekvizitler.Rows[0]["a5"].ToString();
                oSheet.Cells[10, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a6"].ToString();
                oSheet.Cells[11, 2] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainrekvizitler.Rows[0]["a7"].ToString();
                oSheet.Cells[13, 2] = "S. W. I. F. T. Bik: " + MyData.dtmainrekvizitler.Rows[0]["a8"].ToString();

                oSheet.Cells[17, 8] = "Adı / Name: DSMF-nin Mərkəzi Sosial Siğorta və Fərdi Uçot İdarəsi";
                oSheet.Cells[18, 8] = "Hesab № / Acc. №: AZ89CTRE" + Environment.NewLine + "00000000000007018506";
                oSheet.Cells[20, 8] = "VÖEN / Tax İD: 1300115511";

                oSheet.Cells[8, 8] = "Adı / Name: DXA";
                oSheet.Cells[9, 8] = "Kodu / Code: 210005";
                oSheet.Cells[10, 8] = "VÖEN / Tax İD: 1401555071";
                oSheet.Cells[11, 8] = "Müxbir hesab / Corr. Acc.: AZ41NABZ" + Environment.NewLine + "01360100000000003944"; ;
                oSheet.Cells[13, 8] = "S. W. I. F. T. Bik: CTREAZ22";

                MyData.updateCommand("baza.accdb", "UPDATE plateshkanomre SET nomre=" + "'" + txtnomreEmekhaqqi.Text + "'");
                
                MyData.insertCommand("baza.accdb", "insert into plateshkaARXIV (TARİX, №, HARDAN, HARA, [ÖDƏNİŞİN TƏYİNATI], MƏBLƏĞ, VALYUTA) Values ('" + dttarixEmekhaqqi.Text + "', '" + txtnomreEmekhaqqi.Text + "/" + txtnomreEmekhaqqi2.Text + "', '" + cbFromEmekHaqqi.Text + "', 'DSMF-nin Mərkəzi Sosial Siğorta və Fərdi Uçot İdarəsi', 'DSMF-na ayırmalar (10%)" + "', '" + txtmanat3faiz.Text + "', '" + "manat" + "')");
                
                MyData.insertCommand("baza.accdb", "insert into Emeliyyatlar (a1, a2, a3) Values ('" + DateTime.Now + "', 'ÖDƏNİŞ TAPŞIRIĞI - " + "ƏMƏK HAQQI" + " - " + txtmanat3faiz.Text + " - DSMF 10%" + "','" + Environment.MachineName + "')");
                
                oSheet.PrintOut(1, 1, 3);
                oXL.DisplayAlerts = false; 
                oWB.Close(SaveChanges: true);
                oXL.Application.Quit();
                plateshkanomreRefresh();
            }

            //15%--------------------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------------------------------------------
            if (checkBox3.Checked == true)
            {
                try { File.Copy("Plateshka.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\15%.xlsx", true); }
                catch { MessageBox.Show("'Plateshka.xlsx' tapılmadı."); }

                oXL = new Excel.Application();
                oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\15%.xlsx"));
                oSheet = (Excel._Worksheet)oWB.Sheets[1];
                oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                oXL.Visible = false;
                oSheet.Activate();
                oSheet.Range["A1"].Select();
                oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

                if (txtmanatEmekhaqqi2.Text == "") txtmanatEmekhaqqi2.Text = "00";
                if (txtmanatEmekhaqqi.Text == "") txtmanatEmekhaqqi.Text = "0";
                if (txtmanatEmekhaqqi2.Text == "") txtmanatEmekhaqqi2.Text = "00";
                if (txtmanat3faiz.Text == "") txtmanat3faiz.Text = "0";
                if (txtmanatEmekhaqqi6.Text == "") txtmanatEmekhaqqi6.Text = "00";
                if (txtmanat22faiz.Text == "") txtmanat22faiz.Text = "0";

                oSheet.Cells[2, 2] = "ÖDƏNİŞ TAPŞIRIĞI № " + txtnomreEmekhaqqi.Text + "/" + txtnomreEmekhaqqi2.Text;
                oSheet.Cells[4, 2] = dt.Day + " " + a + " " + dt.Year + " - ci il";

                oSheet.Cells[22, 8] = "AZN";
                oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + txtmanat22faiz.Text;
                if (txtmanat22faiz.Text.Length < 2) oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + txtmanat22faiz.Text + ".0" + txtmanatEmekhaqqi6.Text;

                oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: " + txtmanatEmekhaqqi6.Text;
                oSheet.Cells[28, 2] = "D1. Ödənişin təyinatı və əsas / Payment details: DSMF-na ayırmalar - (15%)";
                oSheet.Cells[31, 2] = "D2. Ödənişlə əlaqədar əlavə informasiya / Narrative:";
                oSheet.Cells[35, 2] = "D3. Büdcə təsnifatının kodu: 121111";
                oSheet.Cells[35, 8] = "D4. Büdcə səviyyəsinin kodu: 4";

                MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1='" + cbFromEmekHaqqi.Text + "'");
                MyData.dtmainrekvizitler=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

                string adi = MyData.dtmainrekvizitler.Rows[0][1].ToString();
                try
                {
                    for (int t = 0; t < MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; t++)
                    {
                        if (MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(t, 1) == "(") { adi = MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(0, t); t = MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; }
                    }
                }
                catch { }

                oSheet.Cells[17, 2] = "Adı / Name: " + adi;
                oSheet.Cells[18, 2] = "Hesab № / Acc. №: " + MyData.dtmainrekvizitler.Rows[0]["a2"].ToString();
                oSheet.Cells[20, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a3"].ToString();

                oSheet.Cells[8, 2] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0]["a4"].ToString();
                oSheet.Cells[9, 2] = "Kodu / Code: " + MyData.dtmainrekvizitler.Rows[0]["a5"].ToString();
                oSheet.Cells[10, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a6"].ToString();
                oSheet.Cells[11, 2] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainrekvizitler.Rows[0]["a7"].ToString();
                oSheet.Cells[13, 2] = "S. W. I. F. T. Bik: " + MyData.dtmainrekvizitler.Rows[0]["a8"].ToString();

                oSheet.Cells[17, 8] = "Adı / Name: DSMF-nin Mərkəzi Sosial Siğorta və Fərdi Uçot İdarəsi";
                oSheet.Cells[18, 8] = "Hesab № / Acc. №: AZ89CTRE" + Environment.NewLine + "00000000000007018506";
                oSheet.Cells[20, 8] = "VÖEN / Tax İD: 1300115511";

                oSheet.Cells[8, 8] = "Adı / Name: DXA";
                oSheet.Cells[9, 8] = "Kodu / Code: 210005";
                oSheet.Cells[10, 8] = "VÖEN / Tax İD: 1401555071";
                oSheet.Cells[11, 8] = "Müxbir hesab / Corr. Acc.: AZ41NABZ" + Environment.NewLine + "01360100000000003944"; ;
                oSheet.Cells[13, 8] = "S. W. I. F. T. Bik: CTREAZ22";

                MyData.updateCommand("baza.accdb", "UPDATE plateshkanomre SET nomre=" + "'" + txtnomreEmekhaqqi.Text + "'");
                
                MyData.insertCommand("baza.accdb", "insert into plateshkaARXIV (TARİX, №, HARDAN, HARA, [ÖDƏNİŞİN TƏYİNATI], MƏBLƏĞ, VALYUTA) Values ('" + dttarixEmekhaqqi.Text + "', '" + txtnomreEmekhaqqi.Text + "/" + txtnomreEmekhaqqi2.Text + "', '" + cbFromEmekHaqqi.Text + "', 'DSMF-nin Mərkəzi Sosial Siğorta və Fərdi Uçot İdarəsi', 'DSMF-na ayırmalar (15%)" + "', '" + txtmanat22faiz.Text + "', '" + "manat" + "')");
                
                MyData.insertCommand("baza.accdb", "insert into Emeliyyatlar (a1, a2, a3) Values ('" + DateTime.Now + "', 'ÖDƏNİŞ TAPŞIRIĞI - " + "ƏMƏK HAQQI" + " - " + txtmanat22faiz.Text + " - DSMF 15%" + "','" + Environment.MachineName + "')");
                
                oSheet.PrintOut(1, 1, 3);
                oXL.DisplayAlerts = false; 
                oWB.Close(SaveChanges: true);
                oXL.Application.Quit();
                plateshkanomreRefresh();
            }

            //---Sigorta Eden ucun 05%-------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------------------------------------------

            if (checkBox5.Checked == true)
            {
                try { File.Copy("Plateshka.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Sigorta Eden 05%.xlsx", true); }
                catch { MessageBox.Show("'Plateshka.xlsx' tapılmadı."); }

                oXL = new Excel.Application();
                oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Sigorta Eden 05%.xlsx"));
                oSheet = (Excel._Worksheet)oWB.Sheets[1];
                oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                oXL.Visible = false;
                oSheet.Activate();
                oSheet.Range["A1"].Select();
                oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

                if (txtmanatEmekhaqqi.Text == "") txtmanatEmekhaqqi.Text = "0.00";
                if (txtmanat3faiz.Text == "") txtmanat3faiz.Text = "0.00";
                if (txtmanat22faiz.Text == "") txtmanat22faiz.Text = "0.00";
                if (txtmanatgelirvergisi.Text == "") txtmanatgelirvergisi.Text = "0.00";
                if (txtmanatSigEden05faiz.Text == "") txtmanatSigEden05faiz.Text = "0.00";
                if (txtmanatSigolunan05faiz.Text == "") txtmanatSigolunan05faiz.Text = "0.00";

                oSheet.Cells[2, 2] = "ÖDƏNİŞ TAPŞIRIĞI № " + txtnomreEmekhaqqi.Text + "/" + txtnomreEmekhaqqi2.Text;
                oSheet.Cells[4, 2] = dt.Day+ " " + a + " " +dt.Year + " - ci il";
                oSheet.Cells[22, 8] = "AZN";
                oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + txtmanatSigEden05faiz.Text;
                oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: " + txtmanatEmekhaqqi10.Text;
                oSheet.Cells[28, 2] = "D1. Ödənişin təyinatı və əsas / Payment details: İşəgötürənlərin işsizlikdən siğorta haqqı - (0.5%)";
                oSheet.Cells[31, 2] = "D2. Ödənişlə əlaqədar əlavə informasiya / Narrative:";
                oSheet.Cells[35, 2] = "D3. Büdcə təsnifatının kodu: 123100";
                oSheet.Cells[35, 8] = "D4. Büdcə səviyyəsinin kodu: 4";

                MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1='" + cbFromEmekHaqqi.Text + "'");
                MyData.dtmainrekvizitler=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

                string adi = MyData.dtmainrekvizitler.Rows[0][1].ToString();
                try
                {
                    for (int t = 0; t < MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; t++)
                    {
                        if (MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(t, 1) == "(") { adi = MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(0, t); t = MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; }
                    }
                }
                catch { }

                oSheet.Cells[17, 2] = "Adı / Name: " + adi;
                oSheet.Cells[18, 2] = "Hesab № / Acc. №: " + MyData.dtmainrekvizitler.Rows[0]["a2"].ToString();
                oSheet.Cells[20, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a3"].ToString();

                oSheet.Cells[8, 2] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0]["a4"].ToString();
                oSheet.Cells[9, 2] = "Kodu / Code: " + MyData.dtmainrekvizitler.Rows[0]["a5"].ToString();
                oSheet.Cells[10, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a6"].ToString();
                oSheet.Cells[11, 2] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainrekvizitler.Rows[0]["a7"].ToString();
                oSheet.Cells[13, 2] = "S. W. I. F. T. Bik: " + MyData.dtmainrekvizitler.Rows[0]["a8"].ToString();

                oSheet.Cells[17, 8] = "Adı / Name: DSMF-nin Mərkəzi Sosial Siğorta və Fərdi Uçot İdarəsi";
                oSheet.Cells[18, 8] = "Hesab № / Acc. №: AZ53CTRE" + Environment.NewLine + "00000000000007018572";
                oSheet.Cells[20, 8] = "VÖEN / Tax İD: 1300115511";

                oSheet.Cells[8, 8] = "Adı / Name: DXA";
                oSheet.Cells[9, 8] = "Kodu / Code: 210005";
                oSheet.Cells[10, 8] = "VÖEN / Tax İD: 1401555071";
                oSheet.Cells[11, 8] = "Müxbir hesab / Corr. Acc.: AZ41NABZ" + Environment.NewLine + "01360100000000003944"; ;
                oSheet.Cells[13, 8] = "S. W. I. F. T. Bik: CTREAZ22";

                MyData.updateCommand("baza.accdb", "UPDATE plateshkanomre SET nomre=" + "'" + txtnomreEmekhaqqi.Text + "'");
                
                MyData.insertCommand("baza.accdb", "insert into plateshkaARXIV (TARİX, №, HARDAN, HARA, [ÖDƏNİŞİN TƏYİNATI], MƏBLƏĞ, VALYUTA) Values ('" + dttarixEmekhaqqi.Text + "', '" + txtnomreEmekhaqqi.Text + "/" + txtnomreEmekhaqqi2.Text + "', '" + cbFromEmekHaqqi.Text + "', 'DSMF-nin Mərkəzi Sosial Siğorta və Fərdi Uçot İdarəsi', 'İşəgötürənlərin işsizlikdən siğorta haqqı - (0.5%)" + "', '" + txtmanatSigEden05faiz.Text + "', '" + "manat" + "')");
                
                MyData.insertCommand("baza.accdb", "insert into Emeliyyatlar (a1, a2, a3) Values ('" + DateTime.Now + "', 'ÖDƏNİŞ TAPŞIRIĞI - " + "ƏMƏK HAQQI" + " - " + txtmanatSigEden05faiz.Text + " - İşəgötürənlərin işsizlikdən siğorta haqqı - (0.5%)" + "','" + Environment.MachineName + "')");
                
                oSheet.PrintOut(1, 1, 3);
                oXL.DisplayAlerts = false;
                oWB.Close(SaveChanges: true);
                oXL.Application.Quit();
                plateshkanomreRefresh();
            }

            //---Sigorta Olunan ucun 05%-------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------------------------------------------

            if (checkBox6.Checked == true)
            {
                try { File.Copy("Plateshka.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Sigorta Olunan 05%.xlsx", true); }
                catch { MessageBox.Show("'Plateshka.xlsx' tapılmadı."); }

                oXL = new Excel.Application();
                oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Sigorta Olunan 05%.xlsx"));
                oSheet = (Excel._Worksheet)oWB.Sheets[1];
                oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                oXL.Visible = false;
                oSheet.Activate();
                oSheet.Range["A1"].Select();
                oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

                if (txtmanatEmekhaqqi.Text == "") txtmanatEmekhaqqi.Text = "0.00";
                if (txtmanat3faiz.Text == "") txtmanat3faiz.Text = "0.00";
                if (txtmanat22faiz.Text == "") txtmanat22faiz.Text = "0.00";
                if (txtmanatgelirvergisi.Text == "") txtmanatgelirvergisi.Text = "0.00";
                if (txtmanatSigEden05faiz.Text == "") txtmanatSigEden05faiz.Text = "0.00";
                if (txtmanatSigolunan05faiz.Text == "") txtmanatSigolunan05faiz.Text = "0.00";

                oSheet.Cells[2, 2] = "ÖDƏNİŞ TAPŞIRIĞI № " + txtnomreEmekhaqqi.Text + "/" + txtnomreEmekhaqqi2.Text;
                oSheet.Cells[4, 2] = dt.Day + " " + a + " " + dt.Year + " - ci il";
                oSheet.Cells[22, 8] = "AZN";
                oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + txtmanatSigolunan05faiz.Text;
                oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: " + txtmanatEmekhaqqi12.Text;
                oSheet.Cells[28, 2] = "D1. Ödənişin təyinatı və əsas / Payment details: İşləyənlərin işsizlikdən siğorta haqqı - (0.5%)";
                oSheet.Cells[31, 2] = "D2. Ödənişlə əlaqədar əlavə informasiya / Narrative:";
                oSheet.Cells[35, 2] = "D3. Büdcə təsnifatının kodu: 123200";
                oSheet.Cells[35, 8] = "D4. Büdcə səviyyəsinin kodu: 4";

                MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1='" + cbFromEmekHaqqi.Text + "'");
                MyData.dtmainrekvizitler=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

                string adi = MyData.dtmainrekvizitler.Rows[0][1].ToString();
                try
                {
                    for (int t = 0; t < MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; t++)
                    {
                        if (MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(t, 1) == "(") { adi = MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(0, t); t = MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; }
                    }
                }
                catch { }

                oSheet.Cells[17, 2] = "Adı / Name: " + adi;
                oSheet.Cells[18, 2] = "Hesab № / Acc. №: " + MyData.dtmainrekvizitler.Rows[0]["a2"].ToString();
                oSheet.Cells[20, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a3"].ToString();

                oSheet.Cells[8, 2] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0]["a4"].ToString();
                oSheet.Cells[9, 2] = "Kodu / Code: " + MyData.dtmainrekvizitler.Rows[0]["a5"].ToString();
                oSheet.Cells[10, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a6"].ToString();
                oSheet.Cells[11, 2] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainrekvizitler.Rows[0]["a7"].ToString();
                oSheet.Cells[13, 2] = "S. W. I. F. T. Bik: " + MyData.dtmainrekvizitler.Rows[0]["a8"].ToString();
  
                oSheet.Cells[17, 8] = "Adı / Name: DSMF-nin Mərkəzi Sosial Siğorta və Fərdi Uçot İdarəsi";
                oSheet.Cells[18, 8] = "Hesab № / Acc. №: AZ53CTRE" + Environment.NewLine + "00000000000007018572";
                oSheet.Cells[20, 8] = "VÖEN / Tax İD: 1300115511";

                oSheet.Cells[8, 8] = "Adı / Name: DXA";
                oSheet.Cells[9, 8] = "Kodu / Code: 210005";
                oSheet.Cells[10, 8] = "VÖEN / Tax İD: 1401555071";
                oSheet.Cells[11, 8] = "Müxbir hesab / Corr. Acc.: AZ41NABZ" + Environment.NewLine + "01360100000000003944"; ;
                oSheet.Cells[13, 8] = "S. W. I. F. T. Bik: CTREAZ22";

                MyData.updateCommand("baza.accdb", "UPDATE plateshkanomre SET nomre=" + "'" + txtnomreEmekhaqqi.Text + "'");
                
                MyData.insertCommand("baza.accdb", "insert into plateshkaARXIV (TARİX, №, HARDAN, HARA, [ÖDƏNİŞİN TƏYİNATI], MƏBLƏĞ, VALYUTA) Values ('" + dttarixEmekhaqqi.Text + "', '" + txtnomreEmekhaqqi.Text + "/" + txtnomreEmekhaqqi2.Text + "', '" + cbFromEmekHaqqi.Text + "', 'DSMF-nin Mərkəzi Sosial Siğorta və Fərdi Uçot İdarəsi', 'İşləyənlərin işsizlikdən siğorta haqqı - (0.5%)" + "', '" + txtmanatSigolunan05faiz.Text + "', '" + "manat" + "')");
                
                MyData.insertCommand("baza.accdb", "insert into Emeliyyatlar (a1, a2, a3) Values ('" + DateTime.Now + "', 'ÖDƏNİŞ TAPŞIRIĞI - " + "ƏMƏK HAQQI" + " - " + txtmanatSigolunan05faiz.Text + " - İşləyənlərin işsizlikdən siğorta haqqı - (0.5%)" + "','" + Environment.MachineName + "')");
                
                oSheet.PrintOut(1, 1, 3);
                oXL.DisplayAlerts = false;
                oWB.Close(SaveChanges: true);
                oXL.Application.Quit();
                plateshkanomreRefresh();
            }

            //--Gəlir vergisi--------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------------------------------------------
            if (checkBox4.Checked == true)
            {
                try { File.Copy("Plateshka.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\14%.xlsx", true); }
                catch { MessageBox.Show("'Plateshka.xlsx' tapılmadı."); }

                oXL = new Excel.Application();
                oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\14%.xlsx"));
                oSheet = (Excel._Worksheet)oWB.Sheets[1];
                oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                oXL.Visible = false;
                oSheet.Activate();
                oSheet.Range["A1"].Select();
                oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

                oSheet.Cells[2, 2] = "ÖDƏNİŞ TAPŞIRIĞI № " + txtnomreEmekhaqqi.Text + "/" + txtnomreEmekhaqqi2.Text;
                oSheet.Cells[4, 2] = dt.Day + " " + a + " " + dt.Year + " - ci il";

                oSheet.Cells[22, 8] = "AZN";

                //Emek haqqi ucun
                oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + txtmanatgelirvergisi.Text;
                if (txtmanatgelirvergisi.Text.Length < 2) oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + txtmanatgelirvergisi.Text;

                oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: " + txtmanatEmekhaqqi8.Text;
                oSheet.Cells[28, 2] = "D1. Ödənişin təyinatı və əsas / Payment details: " + Environment.NewLine + textBox8.Text;
                oSheet.Cells[31, 2] = "D2. Ödənişlə əlaqədar əlavə informasiya / Narrative:";
                oSheet.Cells[35, 2] = "D3. Büdcə təsnifatının kodu: 111111";
                oSheet.Cells[35, 8] = "D4. Büdcə səviyyəsinin kodu: 1";

                MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1='" + cbFromEmekHaqqi.Text + "'");
                MyData.dtmainrekvizitler=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

                string adi = MyData.dtmainrekvizitler.Rows[0][1].ToString();
                try
                {
                    for (int t = 0; t < MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; t++)
                    {
                        if (MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(t, 1) == "(") { adi = MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(0, t); t = MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; }
                    }
                }
                catch { }

                oSheet.Cells[17, 2] = "Adı / Name: " + adi;
                oSheet.Cells[18, 2] = "Hesab № / Acc. №: " + MyData.dtmainrekvizitler.Rows[0]["a2"].ToString();
                oSheet.Cells[20, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a3"].ToString();

                oSheet.Cells[8, 2] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0]["a4"].ToString();
                oSheet.Cells[9, 2] = "Kodu / Code: " + MyData.dtmainrekvizitler.Rows[0]["a5"].ToString();
                oSheet.Cells[10, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a6"].ToString();
                oSheet.Cells[11, 2] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainrekvizitler.Rows[0]["a7"].ToString();
                oSheet.Cells[13, 2] = "S. W. I. F. T. Bik: " + MyData.dtmainrekvizitler.Rows[0]["a8"].ToString();

                oSheet.Cells[17, 8] = "Adı / Name: Bakı şəhəri Lokal gəlirlər Departamenti";
                oSheet.Cells[18, 8] = "Hesab № / Acc. №: AZ17CTRE" + Environment.NewLine + "00000000000002117131";
                oSheet.Cells[20, 8] = "VÖEN / Tax İD: 1403006271";

                oSheet.Cells[8, 8]  = "Adı / Name: DXA";
                oSheet.Cells[9, 8]  = "Kodu / Code: 210005";
                oSheet.Cells[10, 8] = "VÖEN / Tax İD: 1401555071";
                oSheet.Cells[11, 8] = "Müxbir hesab / Corr. Acc.: AZ41NABZ" + Environment.NewLine + "01360100000000003944"; ;
                oSheet.Cells[13, 8] = "S. W. I. F. T. Bik: CTREAZ22";
          
                MyData.updateCommand("baza.accdb", "UPDATE plateshkanomre SET nomre=" + "'" + txtnomreEmekhaqqi.Text + "'");
                
                MyData.insertCommand("baza.accdb", "insert into plateshkaARXIV (TARİX, №, HARDAN, HARA, [ÖDƏNİŞİN TƏYİNATI], MƏBLƏĞ, VALYUTA) Values ('" + dttarixEmekhaqqi.Text + "', '" + txtnomreEmekhaqqi.Text + "/" + txtnomreEmekhaqqi2.Text + "', '“AGLizinq” QSC (AGB AZN)', 'Bakı şəhəri Lokal gəlirlər Departamenti" + "', '" + textBox8.Text + "', '" + txtmanatgelirvergisi.Text + "', '" + "manat" + "')");
                
                MyData.insertCommand("baza.accdb", "insert into Emeliyyatlar (a1, a2, a3) Values ('" + DateTime.Now + "', 'ÖDƏNİŞ TAPŞIRIĞI - " + "Bakı şəhəri Lokal gəlirlər Departamenti" + " - " + txtmanatgelirvergisi.Text + " (" + textBox8.Text + ")','" + Environment.MachineName + "')");
                
                oSheet.PrintOut(1, 1, 3);
                oXL.DisplayAlerts = false; 
                oWB.Close(SaveChanges: true);
                oXL.Application.Quit();
                plateshkanomreRefresh();
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            if (txtnomreEmekhaqqi.Enabled == false || txtnomreEmekhaqqi2.Enabled == false) { txtnomreEmekhaqqi2.Enabled = true; txtnomreEmekhaqqi.Enabled = true; return; }
            if (txtnomreEmekhaqqi.Enabled == true || txtnomreEmekhaqqi2.Enabled == true) { txtnomreEmekhaqqi.Enabled = false; txtnomreEmekhaqqi2.Enabled = false; }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            if (dttarixEmekhaqqi.Enabled == false) { dttarixEmekhaqqi.Enabled = true; return; }
            if (dttarixEmekhaqqi.Enabled == true) { dttarixEmekhaqqi.Enabled = false; }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            if (!MyCheck.davamYesNo()) return;

            try
            {
                if (checkBox1.Checked == true && txtmanatEmekhaqqi.Text.Substring(txtmanatEmekhaqqi.Text.Length - 3, 1) != ".")
                {
                    MessageBox.Show("Əmək haqqı hissəsi- Məbləğ düzgün qeyd olunmayıb." + Environment.NewLine + "Nümunə: 123.45"); return;
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); return; }

            try
            {
                if (checkBox2.Checked == true && txtmanat3faiz.Text.Substring(txtmanat3faiz.Text.Length - 3, 1) != ".")
                {
                    MessageBox.Show("10% - Məbləğ düzgün qeyd olunmayıb." + Environment.NewLine + "Nümunə: 123.45"); return;
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); return; }

            try
            {
                if (checkBox3.Checked == true && txtmanat22faiz.Text.Substring(txtmanat22faiz.Text.Length - 3, 1) != ".")
                {
                    MessageBox.Show("15% - Məbləğ düzgün qeyd olunmayıb." + Environment.NewLine + "Nümunə: 123.45"); return;
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); return; }

            try
            {
                if (checkBox4.Checked == true && txtmanatgelirvergisi.Text.Substring(txtmanatgelirvergisi.Text.Length - 3, 1) != ".")
                {
                    MessageBox.Show("Gəlir vergisi - Məbləğ düzgün qeyd olunmayıb." + Environment.NewLine + "Nümunə: 123.45"); return;
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); return; }

            try
            {
                if (checkBox5.Checked == true && txtmanatSigEden05faiz.Text.Substring(txtmanatSigEden05faiz.Text.Length - 3, 1) != ".")
                {
                    MessageBox.Show("Siğorta edən 0.5% - Məbləğ düzgün qeyd olunmayıb." + Environment.NewLine + "Nümunə: 123.45"); return;
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); return; }

            try
            {
                if (checkBox6.Checked == true && txtmanatSigolunan05faiz.Text.Substring(txtmanatSigolunan05faiz.Text.Length - 3, 1) != ".")
                {
                    MessageBox.Show("Siğorta olunan 0.5% - Məbləğ düzgün qeyd olunmayıb." + Environment.NewLine + "Nümunə: 123.45"); return;
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); return; }

            DateTime dt = dttarixEmekhaqqi.Value.Date;
            string a = MyChange.TarixSozle(dt);

            //EMEK HAQQİ--------------------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------------------------------------------
            if (checkBox1.Checked == true)
            {
                try { File.Copy("Plateshka.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Emek haqqi.xlsx", true); }
                catch { MessageBox.Show("'Plateshka.xlsx' tapılmadı."); }

                oXL = new Excel.Application();
                oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Emek haqqi.xlsx"));
                oSheet = (Excel._Worksheet)oWB.Sheets[1];
                oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                oXL.Visible = true;
                oSheet.Activate();
                oSheet.Range["A1"].Select();
                oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

                if (txtmanatEmekhaqqi2.Text == "") txtmanatEmekhaqqi2.Text = "00";
                if (txtmanatEmekhaqqi.Text == "") txtmanatEmekhaqqi.Text = "0";
                if (txtmanatEmekhaqqi2.Text == "") txtmanatEmekhaqqi2.Text = "00";
                if (txtmanat3faiz.Text == "") txtmanat3faiz.Text = "0";
                if (txtmanatEmekhaqqi6.Text == "") txtmanatEmekhaqqi6.Text = "00";
                if (txtmanat22faiz.Text == "") txtmanat22faiz.Text = "0";

                oSheet.Cells[2, 2] = "ÖDƏNİŞ TAPŞIRIĞI № " + txtnomreEmekhaqqi.Text + "/" + txtnomreEmekhaqqi2.Text;
                oSheet.Cells[4, 2] = dt.Day + " " + a + " " + dt.Year + " - ci il";
                oSheet.Cells[22, 8] = "AZN";
                oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + txtmanatEmekhaqqi.Text;
                if (txtmanatEmekhaqqi2.Text.Length < 2) oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + txtmanatEmekhaqqi.Text;

                oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: " + txtmanatEmekhaqqi2.Text;
                oSheet.Cells[28, 2] = "D1. Ödənişin təyinatı və əsas / Payment details: Əmək haqqı";
                oSheet.Cells[31, 2] = "D2. Ödənişlə əlaqədar əlavə informasiya / Narrative:";
                oSheet.Cells[35, 2] = "D3. Büdcə təsnifatının kodu:";
                oSheet.Cells[35, 8] = "D4. Büdcə səviyyəsinin kodu:";

                MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1='" + cbFromEmekHaqqi.Text + "'");
                MyData.dtmainrekvizitler=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

                string adi = MyData.dtmainrekvizitler.Rows[0][1].ToString();
                try
                {
                    for (int t = 0; t < MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; t++)
                    {
                        if (MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(t, 1) == "(") { adi = MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(0, t); t = MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; }
                    }
                }
                catch { }

                oSheet.Cells[17, 2] = "Adı / Name: " + adi;
                oSheet.Cells[18, 2] = "Hesab № / Acc. №: " + MyData.dtmainrekvizitler.Rows[0]["a2"].ToString();
                oSheet.Cells[20, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a3"].ToString();

                oSheet.Cells[8, 2] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0]["a4"].ToString();
                oSheet.Cells[9, 2] = "Kodu / Code: " + MyData.dtmainrekvizitler.Rows[0]["a5"].ToString();
                oSheet.Cells[10, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a6"].ToString();
                oSheet.Cells[11, 2] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainrekvizitler.Rows[0]["a7"].ToString();
                oSheet.Cells[13, 2] = "S. W. I. F. T. Bik: " + MyData.dtmainrekvizitler.Rows[0]["a8"].ToString();
               
                oSheet.Cells[17, 8] = "Adı / Name: “AGLizinq” QSC";
                oSheet.Cells[18, 8] = "Hesab № / Acc. №: AZ30AZEG" + Environment.NewLine + "45013944017950107055";
                oSheet.Cells[20, 8] = "VÖEN / Tax İD: 1300616961";

                oSheet.Cells[8, 8] = "Adı / Name: AGBANK ASC";
                oSheet.Cells[9, 8] = "Kodu / Code: 505817";
                oSheet.Cells[10, 8] = "VÖEN / Tax İD: 9900019651";
                oSheet.Cells[11, 8] = "Müxbir hesab / Corr. Acc.: AZ75NABZ" + Environment.NewLine + "01350100000000017944"; ;
                oSheet.Cells[13, 8] = "S. W. I. F. T. Bik: AZEGAZ22";
            }

            //10%--------------------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------------------------------------------
            if (checkBox2.Checked == true)
            {
                try { File.Copy("Plateshka.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\10%.xlsx", true); }
                catch { MessageBox.Show("'Plateshka.xlsx' tapılmadı."); }

                oXL = new Excel.Application();
                oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\10%.xlsx"));
                oSheet = (Excel._Worksheet)oWB.Sheets[1];
                oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                oXL.Visible = true;
                oSheet.Activate();
                oSheet.Range["A1"].Select();
                oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

                if (txtmanatEmekhaqqi2.Text == "") txtmanatEmekhaqqi2.Text = "00";
                if (txtmanatEmekhaqqi.Text == "") txtmanatEmekhaqqi.Text = "0";
                if (txtmanatEmekhaqqi2.Text == "") txtmanatEmekhaqqi2.Text = "00";
                if (txtmanat3faiz.Text == "") txtmanat3faiz.Text = "0";
                if (txtmanatEmekhaqqi6.Text == "") txtmanatEmekhaqqi6.Text = "00";
                if (txtmanat22faiz.Text == "") txtmanat22faiz.Text = "0";

                oSheet.Cells[2, 2] = "ÖDƏNİŞ TAPŞIRIĞI № " + txtnomreEmekhaqqi.Text + "/" + txtnomreEmekhaqqi2.Text;
                oSheet.Cells[4, 2] = dt.Day + " " + a + " " + dt.Year + " - ci il";
                oSheet.Cells[22, 8] = "AZN";
                oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + txtmanat3faiz.Text;
                if (txtmanatEmekhaqqi2.Text.Length < 2) oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + txtmanat3faiz.Text + ".0" + txtmanatEmekhaqqi2.Text;

                oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: " + txtmanatEmekhaqqi4.Text;
                oSheet.Cells[28, 2] = "D1. Ödənişin təyinatı və əsas / Payment details: DSMF-na ayırmalar - (10%)";
                oSheet.Cells[31, 2] = "D2. Ödənişlə əlaqədar əlavə informasiya / Narrative:";
                oSheet.Cells[35, 2] = "D3. Büdcə təsnifatının kodu: 121211";
                oSheet.Cells[35, 8] = "D4. Büdcə səviyyəsinin kodu: 4";

                MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1='" + cbFromEmekHaqqi.Text + "'");
                MyData.dtmainrekvizitler=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

                string adi = MyData.dtmainrekvizitler.Rows[0][1].ToString();
                try
                {
                    for (int t = 0; t < MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; t++)
                    {
                        if (MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(t, 1) == "(") { adi = MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(0, t); t = MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; }
                    }
                }
                catch { }

                oSheet.Cells[17, 2] = "Adı / Name: " + adi;
                oSheet.Cells[18, 2] = "Hesab № / Acc. №: " + MyData.dtmainrekvizitler.Rows[0]["a2"].ToString();
                oSheet.Cells[20, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a3"].ToString();

                oSheet.Cells[8, 2] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0]["a4"].ToString();
                oSheet.Cells[9, 2] = "Kodu / Code: " + MyData.dtmainrekvizitler.Rows[0]["a5"].ToString();
                oSheet.Cells[10, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a6"].ToString();
                oSheet.Cells[11, 2] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainrekvizitler.Rows[0]["a7"].ToString();
                oSheet.Cells[13, 2] = "S. W. I. F. T. Bik: " + MyData.dtmainrekvizitler.Rows[0]["a8"].ToString();

                oSheet.Cells[17, 8] = "Adı / Name: DSMF-nin Mərkəzi Sosial Siğorta və Fərdi Uçot İdarəsi";
                oSheet.Cells[18, 8] = "Hesab № / Acc. №: AZ89CTRE" + Environment.NewLine + "00000000000007018506";
                oSheet.Cells[20, 8] = "VÖEN / Tax İD: 1300115511";

                oSheet.Cells[8, 8] = "Adı / Name: DXA";
                oSheet.Cells[9, 8] = "Kodu / Code: 210005";
                oSheet.Cells[10, 8] = "VÖEN / Tax İD: 1401555071";
                oSheet.Cells[11, 8] = "Müxbir hesab / Corr. Acc.: AZ41NABZ" + Environment.NewLine + "01360100000000003944"; ;
                oSheet.Cells[13, 8] = "S. W. I. F. T. Bik: CTREAZ22";

            }

            //15%--------------------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------------------------------------------
            if (checkBox3.Checked == true)
            {
                try { File.Copy("Plateshka.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\15%.xlsx", true); }
                catch { MessageBox.Show("'Plateshka.xlsx' tapılmadı."); }

                oXL = new Excel.Application();
                oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\15%.xlsx"));
                oSheet = (Excel._Worksheet)oWB.Sheets[1];
                oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                oXL.Visible = true;
                oSheet.Activate();
                oSheet.Range["A1"].Select();
                oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

                if (txtmanatEmekhaqqi2.Text == "") txtmanatEmekhaqqi2.Text = "00";
                if (txtmanatEmekhaqqi.Text == "") txtmanatEmekhaqqi.Text = "0";
                if (txtmanatEmekhaqqi2.Text == "") txtmanatEmekhaqqi2.Text = "00";
                if (txtmanat3faiz.Text == "") txtmanat3faiz.Text = "0";
                if (txtmanatEmekhaqqi6.Text == "") txtmanatEmekhaqqi6.Text = "00";
                if (txtmanat22faiz.Text == "") txtmanat22faiz.Text = "0";

                oSheet.Cells[2, 2] = "ÖDƏNİŞ TAPŞIRIĞI № " + txtnomreEmekhaqqi.Text + "/" + txtnomreEmekhaqqi2.Text;
                oSheet.Cells[4, 2] = dt.Day + " " + a + " " + dt.Year + " - ci il";

                oSheet.Cells[22, 8] = "AZN";
                oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + txtmanat22faiz.Text;
                if (txtmanat22faiz.Text.Length < 2) oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + txtmanat22faiz.Text + ".0" + txtmanatEmekhaqqi6.Text;

                oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: " + txtmanatEmekhaqqi6.Text;
                oSheet.Cells[28, 2] = "D1. Ödənişin təyinatı və əsas / Payment details: DSMF-na ayırmalar - (15%)";
                oSheet.Cells[31, 2] = "D2. Ödənişlə əlaqədar əlavə informasiya / Narrative:";
                oSheet.Cells[35, 2] = "D3. Büdcə təsnifatının kodu: 121111";
                oSheet.Cells[35, 8] = "D4. Büdcə səviyyəsinin kodu: 4";

                MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1='" + cbFromEmekHaqqi.Text + "'");
                MyData.dtmainrekvizitler=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

                string adi = MyData.dtmainrekvizitler.Rows[0][1].ToString();
                try
                {
                    for (int t = 0; t < MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; t++)
                    {
                        if (MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(t, 1) == "(") { adi = MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(0, t); t = MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; }
                    }
                }
                catch { }

                oSheet.Cells[17, 2] = "Adı / Name: " + adi;
                oSheet.Cells[18, 2] = "Hesab № / Acc. №: " + MyData.dtmainrekvizitler.Rows[0]["a2"].ToString();
                oSheet.Cells[20, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a3"].ToString();

                oSheet.Cells[8, 2] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0]["a4"].ToString();
                oSheet.Cells[9, 2] = "Kodu / Code: " + MyData.dtmainrekvizitler.Rows[0]["a5"].ToString();
                oSheet.Cells[10, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a6"].ToString();
                oSheet.Cells[11, 2] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainrekvizitler.Rows[0]["a7"].ToString();
                oSheet.Cells[13, 2] = "S. W. I. F. T. Bik: " + MyData.dtmainrekvizitler.Rows[0]["a8"].ToString();

                oSheet.Cells[17, 8] = "Adı / Name: DSMF-nin Mərkəzi Sosial Siğorta və Fərdi Uçot İdarəsi";
                oSheet.Cells[18, 8] = "Hesab № / Acc. №: AZ89CTRE" + Environment.NewLine + "00000000000007018506";
                oSheet.Cells[20, 8] = "VÖEN / Tax İD: 1300115511";

                oSheet.Cells[8, 8] = "Adı / Name: DXA";
                oSheet.Cells[9, 8] = "Kodu / Code: 210005";
                oSheet.Cells[10, 8] = "VÖEN / Tax İD: 1401555071";
                oSheet.Cells[11, 8] = "Müxbir hesab / Corr. Acc.: AZ41NABZ" + Environment.NewLine + "01360100000000003944"; ;
                oSheet.Cells[13, 8] = "S. W. I. F. T. Bik: CTREAZ22";

            }

            //---Sigorta Eden ucun 05%-------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------------------------------------------

            if (checkBox5.Checked == true)
            {
                try { File.Copy("Plateshka.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Sigorta Eden 05%.xlsx", true); }
                catch { MessageBox.Show("'Plateshka.xlsx' tapılmadı."); }

                oXL = new Excel.Application();
                oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Sigorta Eden 05%.xlsx"));
                oSheet = (Excel._Worksheet)oWB.Sheets[1];
                oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                oXL.Visible = true;
                oSheet.Activate();
                oSheet.Range["A1"].Select();
                oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

                if (txtmanatEmekhaqqi.Text == "") txtmanatEmekhaqqi.Text = "0.00";
                if (txtmanat3faiz.Text == "") txtmanat3faiz.Text = "0.00";
                if (txtmanat22faiz.Text == "") txtmanat22faiz.Text = "0.00";
                if (txtmanatgelirvergisi.Text == "") txtmanatgelirvergisi.Text = "0.00";
                if (txtmanatSigEden05faiz.Text == "") txtmanatSigEden05faiz.Text = "0.00";
                if (txtmanatSigolunan05faiz.Text == "") txtmanatSigolunan05faiz.Text = "0.00";

                oSheet.Cells[2, 2] = "ÖDƏNİŞ TAPŞIRIĞI № " + txtnomreEmekhaqqi.Text + "/" + txtnomreEmekhaqqi2.Text;
                oSheet.Cells[4, 2] = dt.Day + " " + a + " " + dt.Year + " - ci il";
                oSheet.Cells[22, 8] = "AZN";
                oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + txtmanatSigEden05faiz.Text;
                oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: " + txtmanatEmekhaqqi10.Text;
                oSheet.Cells[28, 2] = "D1. Ödənişin təyinatı və əsas / Payment details: İşəgötürənlərin işsizlikdən siğorta haqqı - (0.5%)";
                oSheet.Cells[31, 2] = "D2. Ödənişlə əlaqədar əlavə informasiya / Narrative:";
                oSheet.Cells[35, 2] = "D3. Büdcə təsnifatının kodu: 123100";
                oSheet.Cells[35, 8] = "D4. Büdcə səviyyəsinin kodu: 4";

                MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1='" + cbFromEmekHaqqi.Text + "'");
                MyData.dtmainrekvizitler=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

                string adi = MyData.dtmainrekvizitler.Rows[0][1].ToString();
                try
                {
                    for (int t = 0; t < MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; t++)
                    {
                        if (MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(t, 1) == "(") { adi = MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(0, t); t = MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; }
                    }
                }
                catch { }

                oSheet.Cells[17, 2] = "Adı / Name: " + adi;
                oSheet.Cells[18, 2] = "Hesab № / Acc. №: " + MyData.dtmainrekvizitler.Rows[0]["a2"].ToString();
                oSheet.Cells[20, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a3"].ToString();

                oSheet.Cells[8, 2] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0]["a4"].ToString();
                oSheet.Cells[9, 2] = "Kodu / Code: " + MyData.dtmainrekvizitler.Rows[0]["a5"].ToString();
                oSheet.Cells[10, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a6"].ToString();
                oSheet.Cells[11, 2] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainrekvizitler.Rows[0]["a7"].ToString();
                oSheet.Cells[13, 2] = "S. W. I. F. T. Bik: " + MyData.dtmainrekvizitler.Rows[0]["a8"].ToString();

                oSheet.Cells[17, 8] = "Adı / Name: DSMF-nin Mərkəzi Sosial Siğorta və Fərdi Uçot İdarəsi";
                oSheet.Cells[18, 8] = "Hesab № / Acc. №: AZ53CTRE" + Environment.NewLine + "00000000000007018572";
                oSheet.Cells[20, 8] = "VÖEN / Tax İD: 1300115511";

                oSheet.Cells[8, 8] = "Adı / Name: DXA";
                oSheet.Cells[9, 8] = "Kodu / Code: 210005";
                oSheet.Cells[10, 8] = "VÖEN / Tax İD: 1401555071";
                oSheet.Cells[11, 8] = "Müxbir hesab / Corr. Acc.: AZ41NABZ" + Environment.NewLine + "01360100000000003944"; ;
                oSheet.Cells[13, 8] = "S. W. I. F. T. Bik: CTREAZ22";
            }

            //---Sigorta Olunan ucun 05%-------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------------------------------------------

            if (checkBox6.Checked == true)
            {
                try { File.Copy("Plateshka.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Sigorta Olunan 05%.xlsx", true); }
                catch { MessageBox.Show("'Plateshka.xlsx' tapılmadı."); }

                oXL = new Excel.Application();
                oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Sigorta Olunan 05%.xlsx"));
                oSheet = (Excel._Worksheet)oWB.Sheets[1];
                oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                oXL.Visible = true;
                oSheet.Activate();
                oSheet.Range["A1"].Select();
                oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

                if (txtmanatEmekhaqqi.Text == "") txtmanatEmekhaqqi.Text = "0.00";
                if (txtmanat3faiz.Text == "") txtmanat3faiz.Text = "0.00";
                if (txtmanat22faiz.Text == "") txtmanat22faiz.Text = "0.00";
                if (txtmanatgelirvergisi.Text == "") txtmanatgelirvergisi.Text = "0.00";
                if (txtmanatSigEden05faiz.Text == "") txtmanatSigEden05faiz.Text = "0.00";
                if (txtmanatSigolunan05faiz.Text == "") txtmanatSigolunan05faiz.Text = "0.00";

                oSheet.Cells[2, 2] = "ÖDƏNİŞ TAPŞIRIĞI № " + txtnomreEmekhaqqi.Text + "/" + txtnomreEmekhaqqi2.Text;
                oSheet.Cells[4, 2] = dt.Day + " " + a + " " + dt.Year + " - ci il";
                oSheet.Cells[22, 8] = "AZN";
                oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + txtmanatSigolunan05faiz.Text;
                oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: " + txtmanatEmekhaqqi12.Text;
                oSheet.Cells[28, 2] = "D1. Ödənişin təyinatı və əsas / Payment details: İşləyənlərin işsizlikdən siğorta haqqı - (0.5%)";
                oSheet.Cells[31, 2] = "D2. Ödənişlə əlaqədar əlavə informasiya / Narrative:";
                oSheet.Cells[35, 2] = "D3. Büdcə təsnifatının kodu: 123200";
                oSheet.Cells[35, 8] = "D4. Büdcə səviyyəsinin kodu: 4";

                MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1='" + cbFromEmekHaqqi.Text + "'");
                MyData.dtmainrekvizitler=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

                string adi = MyData.dtmainrekvizitler.Rows[0][1].ToString();
                try
                {
                    for (int t = 0; t < MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; t++)
                    {
                        if (MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(t, 1) == "(") { adi = MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(0, t); t = MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; }
                    }
                }
                catch { }

                oSheet.Cells[17, 2] = "Adı / Name: " + adi;
                oSheet.Cells[18, 2] = "Hesab № / Acc. №: " + MyData.dtmainrekvizitler.Rows[0]["a2"].ToString();
                oSheet.Cells[20, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a3"].ToString();

                oSheet.Cells[8, 2] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0]["a4"].ToString();
                oSheet.Cells[9, 2] = "Kodu / Code: " + MyData.dtmainrekvizitler.Rows[0]["a5"].ToString();
                oSheet.Cells[10, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a6"].ToString();
                oSheet.Cells[11, 2] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainrekvizitler.Rows[0]["a7"].ToString();
                oSheet.Cells[13, 2] = "S. W. I. F. T. Bik: " + MyData.dtmainrekvizitler.Rows[0]["a8"].ToString();

                oSheet.Cells[17, 8] = "Adı / Name: DSMF-nin Mərkəzi Sosial Siğorta və Fərdi Uçot İdarəsi";
                oSheet.Cells[18, 8] = "Hesab № / Acc. №: AZ53CTRE" + Environment.NewLine + "00000000000007018572";
                oSheet.Cells[20, 8] = "VÖEN / Tax İD: 1300115511";

                oSheet.Cells[8, 8] = "Adı / Name: DXA";
                oSheet.Cells[9, 8] = "Kodu / Code: 210005";
                oSheet.Cells[10, 8] = "VÖEN / Tax İD: 1401555071";
                oSheet.Cells[11, 8] = "Müxbir hesab / Corr. Acc.: AZ41NABZ" + Environment.NewLine + "01360100000000003944"; ;
                oSheet.Cells[13, 8] = "S. W. I. F. T. Bik: CTREAZ22";
            }

            //---Gelir vergisi-------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------------------------------------------
            if (checkBox4.Checked == true)
            {
                try { File.Copy("Plateshka.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\14%.xlsx", true); }
                catch { MessageBox.Show("'Plateshka.xlsx' tapılmadı."); }

                oXL = new Excel.Application();
                oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\14%.xlsx"));
                oSheet = (Excel._Worksheet)oWB.Sheets[1];
                oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                oXL.Visible = true;
                oSheet.Activate();
                oSheet.Range["A1"].Select();
                oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

                if (txtmanatEmekhaqqi2.Text == "") txtmanatEmekhaqqi2.Text = "00";
                if (txtmanatEmekhaqqi.Text == "") txtmanatEmekhaqqi.Text = "0";
                if (txtmanatEmekhaqqi2.Text == "") txtmanatEmekhaqqi2.Text = "00";
                if (txtmanat3faiz.Text == "") txtmanat3faiz.Text = "0";
                if (txtmanatEmekhaqqi6.Text == "") txtmanatEmekhaqqi6.Text = "00";
                if (txtmanat22faiz.Text == "") txtmanat22faiz.Text = "0";
                if (txtmanatEmekhaqqi8.Text == "") txtmanatEmekhaqqi6.Text = "00";
                if (txtmanatgelirvergisi.Text == "") txtmanat22faiz.Text = "0";

                oSheet.Cells[2, 2] = "ÖDƏNİŞ TAPŞIRIĞI № " + txtnomreEmekhaqqi.Text + "/" + txtnomreEmekhaqqi2.Text;
                oSheet.Cells[4, 2] = dt.Day + " " + a + " " + dt.Year + " - ci il";
                oSheet.Cells[22, 8] = "AZN";
                oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + txtmanatgelirvergisi.Text;
                oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: " + txtmanatEmekhaqqi8.Text;
                oSheet.Cells[28, 2] = "D1. Ödənişin təyinatı və əsas / Payment details: " + Environment.NewLine + textBox8.Text;
                oSheet.Cells[31, 2] = "D2. Ödənişlə əlaqədar əlavə informasiya / Narrative:";
                oSheet.Cells[35, 2] = "D3. Büdcə təsnifatının kodu: 111111";
                oSheet.Cells[35, 8] = "D4. Büdcə səviyyəsinin kodu: 1";

                MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1='" + cbFromEmekHaqqi.Text + "'");
                MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

                string adi = MyData.dtmainrekvizitler.Rows[0][1].ToString();
                try
                {
                    for (int t = 0; t < MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; t++)
                    {
                        if (MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(t, 1) == "(") { adi = MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(0, t); t = MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; }
                    }
                }
                catch { }

                oSheet.Cells[17, 2] = "Adı / Name: " + adi;
                oSheet.Cells[18, 2] = "Hesab № / Acc. №: " + MyData.dtmainrekvizitler.Rows[0]["a2"].ToString();
                oSheet.Cells[20, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a3"].ToString();

                oSheet.Cells[8, 2] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0]["a4"].ToString();
                oSheet.Cells[9, 2] = "Kodu / Code: " + MyData.dtmainrekvizitler.Rows[0]["a5"].ToString();
                oSheet.Cells[10, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a6"].ToString();
                oSheet.Cells[11, 2] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainrekvizitler.Rows[0]["a7"].ToString();
                oSheet.Cells[13, 2] = "S. W. I. F. T. Bik: " + MyData.dtmainrekvizitler.Rows[0]["a8"].ToString();

                oSheet.Cells[17, 8] = "Adı / Name: Bakı şəhəri Lokal gəlirlər Departamenti";
                oSheet.Cells[18, 8] = "Hesab № / Acc. №: AZ17CTRE" + Environment.NewLine + "00000000000002117131";
                oSheet.Cells[20, 8] = "VÖEN / Tax İD: 1403006271";

                oSheet.Cells[8, 8] = "Adı / Name: DXA";
                oSheet.Cells[9, 8] = "Kodu / Code: 210005";
                oSheet.Cells[10, 8] = "VÖEN / Tax İD: 1401555071";
                oSheet.Cells[11, 8] = "Müxbir hesab / Corr. Acc.: AZ41NABZ" + Environment.NewLine + "01360100000000003944"; ;
                oSheet.Cells[13, 8] = "S. W. I. F. T. Bik: CTREAZ22";

            }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            if (!MyCheck.davamYesNo()) return;

            try
            {
                if (txtmanat.Text.Substring(txtmanat.Text.Length - 3, 1) != ".")
                {
                    MessageBox.Show("Məbləğ düzgün qeyd olunmayıb." + Environment.NewLine + "Nümunə: 123.45"); return;
                }
            }
            catch { MessageBox.Show("Məbləğ düzgün qeyd olunmayıb." + Environment.NewLine + "Nümunə: 123.45"); return; }


            try { File.Copy("Plateshka.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Odenis Tapsirigi - " + txtmanat.Text + ".xlsx", true); }
            catch { MessageBox.Show("'Plateshka.xlsx' tapılmadı."); }

            try
            {

                if (comboBox2.Text == "") { MessageBox.Show("Benefisiar (Alan) müştəri seçilməyib.."); return; }

                DateTime dt = dttarix.Value.Date;
                string a = MyChange.TarixSozle(dt);
                
                //Get a new workbook.
                oXL = new Excel.Application();
                oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Odenis Tapsirigi - " + txtmanat.Text + ".xlsx"));
                oSheet = (Excel._Worksheet)oWB.Sheets[1];
                oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                oXL.Visible = true;
                oSheet.Activate();
                oSheet.Range["A1"].Select();
                oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

                if (txtmanat.Text == "") txtmanat.Text = "0";

                oSheet.Cells[2, 2] = "ÖDƏNİŞ TAPŞIRIĞI № " + txtnomresi.Text + "/" + txtnomresi2.Text;
                oSheet.Cells[4, 2] = dt.Day + " " + a + " " + dt.Year + " - ci il";
                if (comboBox3.Text == "manat") oSheet.Cells[22, 8] = "AZN";
                if (comboBox3.Text == "dollar") oSheet.Cells[22, 8] = "USD";
                if (comboBox3.Text == "avro") oSheet.Cells[22, 8] = "EUR";

                reqemler();
                oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + txtmanat.Text;

                oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: " + txt2.Text;
                if (comboBox3.Text == "dollar") oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: " + txt2.Text.Substring(0, txt2.Text.Length - 15) + ", 00 " + "ABŞ dolları";
                oSheet.Cells[28, 2] = "D1. Ödənişin təyinatı və əsas / Payment details: " + txtteyinat.Text;
                oSheet.Cells[31, 2] = "D2. Ödənişlə əlaqədar əlavə informasiya / Narrative:";

                MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1=" + "'" + comboBox1.Text + "'");
                MyData.dtmainrekvizitler=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

                string adi = MyData.dtmainrekvizitler.Rows[0][1].ToString();
                try
                {
                    for (int t = 0; t < MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; t++)
                    {
                        if (MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(t, 1) == "(") { adi = MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(0, t); t = MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; }
                    }
                }
                catch { }
                oSheet.Cells[17, 2] = "Adı / Name: " + adi;

                oSheet.Cells[18, 2] = "Hesab № / Acc. №: " + MyData.dtmainrekvizitler.Rows[0][2].ToString();
                oSheet.Cells[20, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0][3].ToString();

                oSheet.Cells[8, 2] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0][4].ToString();
                oSheet.Cells[9, 2] = "Kodu / Code: " + MyData.dtmainrekvizitler.Rows[0][5].ToString();
                oSheet.Cells[10, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0][6].ToString();
                oSheet.Cells[11, 2] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainrekvizitler.Rows[0][7].ToString();
                if (comboBox2.Text == "“AGLizinq” QSC (UNİ USD Daxili)") oSheet.Cells[11, 2] = "Müxbir hesab / Corr. Acc.:";
                oSheet.Cells[13, 2] = "S. W. I. F. T. Bik: " + MyData.dtmainrekvizitler.Rows[0][8].ToString();

                MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1=" + "'" + comboBox2.Text + "'");
                MyData.dtmainrekvizitler=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

                adi = MyData.dtmainrekvizitler.Rows[0][1].ToString();
                try
                {
                    for (int t = 0; t < MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; t++)
                    {
                        if (MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(t, 1) == "(") { adi = MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(0, t); t = MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; }
                    }
                }
                catch { }
                oSheet.Cells[17, 8] = "Adı / Name: " + adi;

                oSheet.Cells[18, 8] = "Hesab № / Acc. №: " + MyData.dtmainrekvizitler.Rows[0][2].ToString();
                oSheet.Cells[20, 8] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0][3].ToString();
                if (MyData.dtmainrekvizitler.Rows[0][9].ToString() != "") oSheet.Cells[20, 8] = "Ş/V: " + MyData.dtmainrekvizitler.Rows[0][9].ToString();

                oSheet.Cells[8, 8] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0][4].ToString();
                oSheet.Cells[9, 8] = "Kodu / Code: " + MyData.dtmainrekvizitler.Rows[0][5].ToString();
                oSheet.Cells[10, 8] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0][6].ToString();
                oSheet.Cells[11, 8] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainrekvizitler.Rows[0][7].ToString();
                oSheet.Cells[13, 8] = "S. W. I. F. T. Bik: " + MyData.dtmainrekvizitler.Rows[0][8].ToString();
            }
            catch { }
        }

        private void button20_Click(object sender, EventArgs e)
        {
            if (!MyCheck.davamYesNo()) return;

            try
            {
                if (txtedvmanat.Text.Substring(txtedvmanat.Text.Length - 3, 1) != ".")
                {
                    MessageBox.Show("Məbləğ düzgün qeyd olunmayıb." + Environment.NewLine + "Nümunə: 123.45"); return;
                }
            }
            catch { MessageBox.Show("Məbləğ düzgün qeyd olunmayıb." + Environment.NewLine + "Nümunə: 123.45"); return; }

            try { File.Copy("EDV.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Odenis Tapsirigi - " + txtedvmanat.Text  +  ".xlsx", true); }
            catch { MessageBox.Show("'EDV.xlsx' tapılmadı."); }

            try
            {

                if (cbedv2.Text == "") { MessageBox.Show("Benefisiar (Alan) müştəri seçilməyib.."); return; }

                DateTime dt = dtedvtarix.Value.Date;
                string a = MyChange.TarixSozle(dt);
              
                //Get a new workbook.
                oXL = new Excel.Application();
                oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Odenis Tapsirigi - " + txtedvmanat.Text +  ".xlsx"));
                oSheet = (Excel._Worksheet)oWB.Sheets[1];
                oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                oXL.Visible = true;
                oSheet.Activate();
                oSheet.Range["A1"].Select();
                oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

                if (txtedvmanat.Text == "") txtmanat.Text = "0";

                oSheet.Cells[3, 3] = "'A' formatlı ödəniş tapşırığı № " + txtedvnomre.Text + "/" + txtedvnomre2.Text;
                oSheet.Cells[4, 3] = dt.Day + " " + a + " " + dt.Year + " - ci il";

                if (comboBox6.Text == "manat") oSheet.Cells[22, 9] = "AZN";
                if (comboBox6.Text == "dollar") oSheet.Cells[22, 9] = "USD";
                if (comboBox6.Text == "avro") oSheet.Cells[22, 9] = "EUR";

                reqemler2();
                oSheet.Cells[24, 3] = "Məbləğ rəqəmlə: " + txtedvmanat.Text ;

                oSheet.Cells[25, 3] = "Məbləğ yazı ilə / İn words: " + txtedvherfle.Text;
                if (comboBox3.Text == "dollar") oSheet.Cells[25, 3] = "Məbləğ yazı ilə / İn words: " + txtedvherfle.Text.Substring(0, txtedvherfle.Text.Length - 15) + ", 00 " + "ABŞ dolları";
                oSheet.Cells[26, 3] = "D1. Ödənişin təyinatı və əsas / Payment details: " + txtedvteyinat.Text;

                MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1=" + "'" + cbedv1.Text + "'");
                MyData.dtmainrekvizitler=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

                string adi = MyData.dtmainrekvizitler.Rows[0][1].ToString();
                try
                {
                    for (int t = 0; t < MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; t++)
                    {
                        if (MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(t, 1) == "(") { adi = MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(0, t); t = MyData.dtmainrekvizitler.Rows[0][1].ToString().Length; }
                    }
                }
                catch { }
                oSheet.Cells[17, 3] = "Adı / Name: " + adi;

                oSheet.Cells[18, 3] = "Hesab № / Acc. №: " + MyData.dtmainrekvizitler.Rows[0][2].ToString();
                oSheet.Cells[20, 3] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0][3].ToString();
                oSheet.Cells[8, 3] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0][4].ToString();
                oSheet.Cells[9, 3] = "Kodu / Code: " + MyData.dtmainrekvizitler.Rows[0][5].ToString();
                oSheet.Cells[10, 3] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0][6].ToString();
                oSheet.Cells[11, 3] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainrekvizitler.Rows[0][7].ToString();
                oSheet.Cells[13, 3] = "S. W. I. F. T. Bik: " + MyData.dtmainrekvizitler.Rows[0][8].ToString();

                MyData.selectCommand("baza.accdb", "Select * From plateshkaEDV WHERE a1='" + cbedv2.Text + "'");
                MyData.dtmainedv=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainedv);

                adi = MyData.dtmainedv.Rows[0]["a1"].ToString();
                try
                {
                    for (int t = 0; t < MyData.dtmainedv.Rows[0]["a1"].ToString().Length; t++)
                    {
                        if (MyData.dtmainedv.Rows[0]["a1"].ToString().Substring(t, 1) == "(") { adi = MyData.dtmainedv.Rows[0]["a1"].ToString().Substring(0, t); t = MyData.dtmainedv.Rows[0]["a1"].ToString().Length; }
                    }
                }
                catch { }
                oSheet.Cells[17, 9] = "Adı / Name: " + adi;

                oSheet.Cells[18, 9] = "Hesab № / Acc. №: " + MyData.dtmainedv.Rows[0]["a2"].ToString();
                oSheet.Cells[20, 9] = "VÖEN / Tax İD: " + MyData.dtmainedv.Rows[0]["a3"].ToString();
                oSheet.Cells[8, 9] = "Adı / Name: " + MyData.dtmainedv.Rows[0]["a4"].ToString();
                oSheet.Cells[9, 9] = "Kodu / Code: " + MyData.dtmainedv.Rows[0]["a5"].ToString();
                oSheet.Cells[10, 9] = "VÖEN / Tax İD: " + MyData.dtmainedv.Rows[0]["a6"].ToString();
                oSheet.Cells[11, 9] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainedv.Rows[0]["a7"].ToString();
                oSheet.Cells[13, 9] = "S. W. I. F. T. Bik: " + MyData.dtmainedv.Rows[0]["a8"].ToString();
                oSheet.Cells[31, "F"] = MyData.dtmainedv.Rows[0]["a9"].ToString();
                oSheet.Cells[31, "L"] = MyData.dtmainedv.Rows[0]["a10"].ToString();
            }
            catch { }
        }

        private void button22_Click(object sender, EventArgs e)
        {
            try
            {
                MyData.updateCommand("baza.accdb", "UPDATE rekvizitler SET "
                                                                                     + "a1 ='" + txtmusteriadi.Text + "',"
                                                                                     + "a2 ='" + txthesab1.Text + " " + txthesab2.Text + "',"
                                                                                     + "a3 ='" + txtmusterivoen.Text + "',"
                                                                                     + "a4 ='" + txtbankadi.Text + "',"
                                                                                     + "a5 ='" + txtbankkodu.Text + "',"
                                                                                     + "a6 ='" + txtbankvoen.Text + "',"
                                                                                     + "a7 ='" + txtmuxhesab1.Text + " " + txtmuxhesab2.Text + "',"
                                                                                     + "a8 ='" + txtsvift.Text + "',"
                                                                                     + "a9 ='" + txtsexsiyyet.Text + "'"
                                                                                     + " WHERE a1 Like '" + txtmusteriadi.Text + "'");

                
                

                MessageBox.Show("Yadda saxlanıldı.");
            }
            catch { MessageBox.Show("Əməliyyatda səhv var." + Environment.NewLine + " - Bu adda müştəri yaddaşda yoxdur"); }

            try
            {
                MyData.updateCommand("baza.accdb", "UPDATE plateshkateyinat SET "
                                                                                     + "a2 ='" + txtesas.Text + "'"
                                                                                     + " WHERE a1 Like '" + txtmusteriadi.Text + "'");

                
                
            }
            catch { }
        }

        private void ödənişTapşırığıToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try { File.Copy("Plateshka.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Odenis Tapsirigi.xlsx", true); }
            catch { MessageBox.Show("'Plateshka.xlsx' tapılmadı."); }

            DateTime dt = Convert.ToDateTime(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["TARİX"].Value);
            string a = MyChange.TarixSozle(dt);
            
            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Odenis Tapsirigi.xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

            oSheet.Cells[2, 2] = "ÖDƏNİŞ TAPŞIRIĞI № " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["No"].Value.ToString();
            oSheet.Cells[4, 2] = dt.Day + " " + a + " " + dt.Year + " - ci il";
            if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["VALYUTA"].Value.ToString() == "manat")
            {
                oSheet.Cells[22, 8] = "AZN";
                try
                {
                    oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: " + MyChange.ReqemToMetn(Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["MƏBLƏĞ"].Value));
                }
                catch { }
            }
            else if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["VALYUTA"].Value.ToString() == "dollar")
            {
                oSheet.Cells[22, 8] = "USD";
                try
                {
                    oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: " + MyChange.ReqemToMetnValyuta(Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["MƏBLƏĞ"].Value),"dollar","sent");
                }
                catch { }
            }
            oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["MƏBLƏĞ"].Value.ToString();

            oSheet.Cells[28, 2] = "D1. Ödənişin təyinatı və əsas / Payment details: " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["TƏYİNAT"].Value.ToString();

            MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1=" + "'" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["HARDAN"].Value.ToString() + "'");
            MyData.dtmainrekvizitler=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

            oSheet.Cells[17, 2] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0][1].ToString();
            try { if (MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(1, 8) == "AGLizinq") oSheet.Cells[17, 2] = "Adı / Name: " + "“AGLizinq” QSC"; }
            catch { }
            oSheet.Cells[18, 2] = "Hesab № / Acc. №: " + MyData.dtmainrekvizitler.Rows[0][2].ToString();
            oSheet.Cells[20, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0][3].ToString();

            oSheet.Cells[8, 2] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0][4].ToString();
            oSheet.Cells[9, 2] = "Kodu / Code: " + MyData.dtmainrekvizitler.Rows[0][5].ToString();
            oSheet.Cells[10, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0][6].ToString();
            oSheet.Cells[11, 2] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainrekvizitler.Rows[0][7].ToString();
            oSheet.Cells[13, 2] = "S. W. I. F. T. Bik: " + MyData.dtmainrekvizitler.Rows[0][8].ToString();
      
            MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["HARA"].Value.ToString() + "'");
            MyData.dtmainrekvizitler=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

            //UNI USD Daxili ucun muxbir hesabin silinmesi. Odeyen bankin muxbir hesabina ehtiyac olmur
            try { if (MyData.dtmainrekvizitler.Rows[0][1].ToString() == "“AGLizinq” QSC (UNİ USD Daxili)") oSheet.Cells[11, 2] = "Müxbir hesab / Corr. Acc.:"; }
            catch { }

            oSheet.Cells[17, 8] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0][1].ToString();
            try { if (MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(1, 8) == "AGLizinq") oSheet.Cells[17, 8] = "Adı / Name: " + "“AGLizinq” QSC"; }
            catch { }
            oSheet.Cells[18, 8] = "Hesab № / Acc. №: " + MyData.dtmainrekvizitler.Rows[0][2].ToString();
            oSheet.Cells[20, 8] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0][3].ToString();
            if (MyData.dtmainrekvizitler.Rows[0][9].ToString() != "") oSheet.Cells[20, 8] = "Ş/V: " + MyData.dtmainrekvizitler.Rows[0][9].ToString();

            oSheet.Cells[8, 8] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0][4].ToString();
            oSheet.Cells[9, 8] = "Kodu / Code: " + MyData.dtmainrekvizitler.Rows[0][5].ToString();
            oSheet.Cells[10, 8] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0][6].ToString();
            oSheet.Cells[11, 8] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainrekvizitler.Rows[0][7].ToString();
            oSheet.Cells[13, 8] = "S. W. I. F. T. Bik: " + MyData.dtmainrekvizitler.Rows[0][8].ToString();
            //////---------------------------------------------------------------------------------------

            oXL.DisplayAlerts = false;
            oWB.Save();
        }

        private void əDZVToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try { File.Copy("EDV.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\EDV.xlsx", true); }
            catch { MessageBox.Show("'EDV.xlsx' tapılmadı."); }

            DateTime dt =Convert.ToDateTime(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["TARİX"].Value);
            string a = MyChange.TarixSozle(dt);
            
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\EDV.xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

            oSheet.Cells[3, 3] = "'A' formatlı ödəniş tapşırığı № " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["No"].Value.ToString();
            oSheet.Cells[4, 3] = dt.Day + " " + a + " " + dt.Year + "-ci il";

            if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["VALYUTA"].Value.ToString() == "manat")
            {
                oSheet.Cells[22, 8] = "AZN";
                try
                {
                    oSheet.Cells[25, 3] = "Məbləğ yazı ilə / İn words: " + MyChange.ReqemToMetn(Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["MƏBLƏĞ"].Value));
                }
                catch { }
            }
            else if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["VALYUTA"].Value.ToString() == "dollar")
            {
                oSheet.Cells[22, 8] = "USD";
                try
                {
                    oSheet.Cells[25, 3] = "Məbləğ yazı ilə / İn words: " + MyChange.ReqemToMetnValyuta(Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["MƏBLƏĞ"].Value), "dollar", "sent");
                }
                catch { }
            }
            oSheet.Cells[24, 3] = "Məbləğ rəqəmlə: " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["MƏBLƏĞ"].Value.ToString();
            
            oSheet.Cells[26, 3] = "D1. Ödənişin təyinatı və əsas / Payment details: " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["TƏYİNAT"].Value.ToString();

            MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1 Like '" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["HARDAN"].Value.ToString() + "'");
            MyData.dtmainrekvizitler=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

            oSheet.Cells[17, 3] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0][1].ToString();
            try { if (MyData.dtmainrekvizitler.Rows[0][1].ToString().Substring(1, 8) == "AGLizinq") oSheet.Cells[17, 3] = "Adı / Name: " + "“AGLizinq” QSC"; }
            catch { }
            oSheet.Cells[18, 3] = "Hesab № / Acc. №: " + MyData.dtmainrekvizitler.Rows[0][2].ToString();
            oSheet.Cells[20, 3] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0][3].ToString();

            oSheet.Cells[8, 3] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0][4].ToString();
            oSheet.Cells[9, 3] = "Kodu / Code: " + MyData.dtmainrekvizitler.Rows[0][5].ToString();
            oSheet.Cells[10, 3] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0][6].ToString();
            oSheet.Cells[11, 3] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainrekvizitler.Rows[0][7].ToString();
            oSheet.Cells[13, 3] = "S. W. I. F. T. Bik: " + MyData.dtmainrekvizitler.Rows[0][8].ToString();
           
            MyData.selectCommand("baza.accdb", "Select * From plateshkaEDV WHERE a1 Like '" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["HARA"].Value.ToString() + "'");
            MyData.dtmainedv=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainedv);

            oSheet.Cells[17, 9] = "Adı / Name: " + MyData.dtmainedv.Rows[0]["a1"].ToString();
            try { if (MyData.dtmainedv.Rows[0]["a1"].ToString().Substring(1, 8) == "AGLizinq") oSheet.Cells[17, 9] = "Adı / Name: " + "“AGLizinq” QSC"; }
            catch { }
            oSheet.Cells[18, 9] = "Hesab № / Acc. №: " + MyData.dtmainedv.Rows[0]["a2"].ToString();
            oSheet.Cells[20, 9] = "VÖEN / Tax İD: " + MyData.dtmainedv.Rows[0]["a3"].ToString();

            oSheet.Cells[8, 9] = "Adı / Name: " + MyData.dtmainedv.Rows[0]["a4"].ToString();
            oSheet.Cells[9, 9] = "Kodu / Code: " + MyData.dtmainedv.Rows[0]["a5"].ToString();
            oSheet.Cells[10, 9] = "VÖEN / Tax İD: " + MyData.dtmainedv.Rows[0]["a6"].ToString();
            oSheet.Cells[11, 9] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainedv.Rows[0]["a7"].ToString();
            oSheet.Cells[13, 9] = "S. W. I. F. T. Bik: " + MyData.dtmainedv.Rows[0]["a8"].ToString();
            oSheet.Cells[31, "F"] = MyData.dtmainedv.Rows[0]["a9"].ToString();
            oSheet.Cells[31, "L"] = MyData.dtmainedv.Rows[0]["a10"].ToString();
            //////---------------------------------------------------------------------------------------
            oXL.DisplayAlerts = false;
            oWB.Save();
        }

        private void label14_Click(object sender, EventArgs e)
        {
            button22.Visible = true;
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            try { File.Copy("Plateshka.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\15%.xlsx", true); }
            catch { MessageBox.Show("'Plateshka.xlsx' tapılmadı."); }

            DateTime dt =Convert.ToDateTime(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["TARİX"].Value);
            string a = MyChange.TarixSozle(dt);
            
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\15%.xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

            oSheet.Cells[2, 2] = "ÖDƏNİŞ TAPŞIRIĞI № " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["No"].Value;
            oSheet.Cells[4, 2] = dt.Day + " " + a + " " + dt.Year + "- ci il";
            oSheet.Cells[22, 8] = "AZN";
            oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["MƏBLƏĞ"].Value;
            oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: ";

            try
            {
                oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: " + MyChange.ReqemToMetn(Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["MƏBLƏĞ"].Value));
            }
            catch { }

            oSheet.Cells[28, 2] = "D1. Ödənişin təyinatı və əsas / Payment details: DSMF-na ayırmalar - (15%)";
            oSheet.Cells[31, 2] = "D2. Ödənişlə əlaqədar əlavə informasiya / Narrative:";
            oSheet.Cells[35, 2] = "D3. Büdcə təsnifatının kodu: 121111";
            oSheet.Cells[35, 8] = "D4. Büdcə səviyyəsinin kodu: 4";

            MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["HARDAN"].Value + "'");
            MyData.dtmainrekvizitler=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

            oSheet.Cells[17, 2] = "Adı / Name: “AGLizinq” QSC";
            oSheet.Cells[18, 2] = "Hesab № / Acc. №: " + MyData.dtmainrekvizitler.Rows[0]["a2"].ToString();
            oSheet.Cells[20, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a3"].ToString();

            oSheet.Cells[8, 2] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0]["a4"].ToString();
            oSheet.Cells[9, 2] = "Kodu / Code: " + MyData.dtmainrekvizitler.Rows[0]["a5"].ToString();
            oSheet.Cells[10, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a6"].ToString();
            oSheet.Cells[11, 2] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainrekvizitler.Rows[0]["a7"].ToString();
            oSheet.Cells[13, 2] = "S. W. I. F. T. Bik: " + MyData.dtmainrekvizitler.Rows[0]["a8"].ToString();
            //////---------------------------------------------------------------------------------------

            ///////-------------------------Benefisiar alan bank ve musterinin rekvizitleri ucun

            if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["HARA"].Value.ToString() == "BHŞİD")
            {
                oSheet.Cells[17, 8] = "Adı / Name: BHŞİD";
                oSheet.Cells[18, 8] = "Hesab № / Acc. №: AZ35CTRE" + Environment.NewLine + "00000000000007018508";
                oSheet.Cells[20, 8] = "VÖEN / Tax İD: 1600351421";

                oSheet.Cells[8, 8] = "Adı / Name: DXA";
                oSheet.Cells[9, 8] = "Kodu / Code: 210005";
                oSheet.Cells[10, 8] = "VÖEN / Tax İD: 1401555071";
                oSheet.Cells[11, 8] = "Müxbir hesab / Corr. Acc.: AZ41NABZ" + Environment.NewLine + "01360100000000003944"; ;
                oSheet.Cells[13, 8] = "S. W. I. F. T. Bik: CTREAZ22";
            }
            else
            {
                oSheet.Cells[17, 8] = "Adı / Name: DSMF-nin Mərkəzi Sosial Siğorta və Fərdi Uçot İdarəsi";
                oSheet.Cells[18, 8] = "Hesab № / Acc. №: AZ89CTRE" + Environment.NewLine + "00000000000007018506";
                oSheet.Cells[20, 8] = "VÖEN / Tax İD: 1300115511";

                oSheet.Cells[8, 8] = "Adı / Name: DXA";
                oSheet.Cells[9, 8] = "Kodu / Code: 210005";
                oSheet.Cells[10, 8] = "VÖEN / Tax İD: 1401555071";
                oSheet.Cells[11, 8] = "Müxbir hesab / Corr. Acc.: AZ41NABZ" + Environment.NewLine + "01360100000000003944"; ;
                oSheet.Cells[13, 8] = "S. W. I. F. T. Bik: CTREAZ22";
            }
            //////---------------------------------------------------------------------------------------

            //oXL.DisplayAlerts = false;
            //oWB.Close(SaveChanges: true);
            //oXL.Application.Quit();
        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            try { File.Copy("Plateshka.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\10%.xlsx", true); }
            catch { MessageBox.Show("'Plateshka.xlsx' tapılmadı."); }

            DateTime dt = Convert.ToDateTime(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["TARİX"].Value);
            string a = MyChange.TarixSozle(dt);

            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\10%.xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

            oSheet.Cells[2, 2] = "ÖDƏNİŞ TAPŞIRIĞI № " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["No"].Value;
            oSheet.Cells[4, 2] = dt.Day + " " + a + " " + dt.Year + "-ci il";
            oSheet.Cells[22, 8] = "AZN";
            //Emek haqqi ucun
            oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["MƏBLƏĞ"].Value;
            oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: ";

            try
            {
                oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: " + MyChange.ReqemToMetn(Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["MƏBLƏĞ"].Value));
            }
            catch { }

            oSheet.Cells[28, 2] = "D1. Ödənişin təyinatı və əsas / Payment details: DSMF-na ayırmalar - (10%)";
            oSheet.Cells[31, 2] = "D2. Ödənişlə əlaqədar əlavə informasiya / Narrative:";
            oSheet.Cells[35, 2] = "D3. Büdcə təsnifatının kodu: 121211";
            oSheet.Cells[35, 8] = "D4. Büdcə səviyyəsinin kodu: 4";

            MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["HARDAN"].Value + "'");
            MyData.dtmainrekvizitler=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

            oSheet.Cells[17, 2] = "Adı / Name: “AGLizinq” QSC";
            oSheet.Cells[18, 2] = "Hesab № / Acc. №: " + MyData.dtmainrekvizitler.Rows[0]["a2"].ToString();
            oSheet.Cells[20, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a3"].ToString();

            oSheet.Cells[8, 2] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0]["a4"].ToString();
            oSheet.Cells[9, 2] = "Kodu / Code: " + MyData.dtmainrekvizitler.Rows[0]["a5"].ToString();
            oSheet.Cells[10, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a6"].ToString();
            oSheet.Cells[11, 2] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainrekvizitler.Rows[0]["a7"].ToString();
            oSheet.Cells[13, 2] = "S. W. I. F. T. Bik: " + MyData.dtmainrekvizitler.Rows[0]["a8"].ToString();

            if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["HARA"].Value.ToString() == "BHŞİD")
            {
                oSheet.Cells[17, 8] = "Adı / Name: BHŞİD";
                oSheet.Cells[18, 8] = "Hesab № / Acc. №: AZ35CTRE" + Environment.NewLine + "00000000000007018508";
                oSheet.Cells[20, 8] = "VÖEN / Tax İD: 1600351421";

                oSheet.Cells[8, 8] = "Adı / Name: DXA";
                oSheet.Cells[9, 8] = "Kodu / Code: 210005";
                oSheet.Cells[10, 8] = "VÖEN / Tax İD: 1401555071";
                oSheet.Cells[11, 8] = "Müxbir hesab / Corr. Acc.: AZ41NABZ" + Environment.NewLine + "01360100000000003944"; ;
                oSheet.Cells[13, 8] = "S. W. I. F. T. Bik: CTREAZ22";
            }
            else
            {
                oSheet.Cells[17, 8] = "Adı / Name: DSMF-nin Mərkəzi Sosial Siğorta və Fərdi Uçot İdarəsi";
                oSheet.Cells[18, 8] = "Hesab № / Acc. №: AZ89CTRE" + Environment.NewLine + "00000000000007018506";
                oSheet.Cells[20, 8] = "VÖEN / Tax İD: 1300115511";

                oSheet.Cells[8, 8] = "Adı / Name: DXA";
                oSheet.Cells[9, 8] = "Kodu / Code: 210005";
                oSheet.Cells[10, 8] = "VÖEN / Tax İD: 1401555071";
                oSheet.Cells[11, 8] = "Müxbir hesab / Corr. Acc.: AZ41NABZ" + Environment.NewLine + "01360100000000003944"; ;
                oSheet.Cells[13, 8] = "S. W. I. F. T. Bik: CTREAZ22";
            }
        }

        private void əməkHaqqıToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try { File.Copy("Plateshka.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Emek haqqi.xlsx", true); }
            catch { MessageBox.Show("'Plateshka.xlsx' tapılmadı."); }

            DateTime dt =Convert.ToDateTime(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["TARİX"].Value);
            string a = MyChange.TarixSozle(dt);

            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Emek haqqi.xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

            oSheet.Cells[2, 2] = "ÖDƏNİŞ TAPŞIRIĞI № " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["No"].Value;
            oSheet.Cells[4, 2] = dt.Day + " " + a + " " + dt.Year + "-ci il";
            oSheet.Cells[22, 8] = "AZN";
            //Emek haqqi ucun
            oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["MƏBLƏĞ"].Value;
            oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: ";


            try
            {
                oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: " + MyChange.ReqemToMetn(Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["MƏBLƏĞ"].Value));
            }
            catch { }

            oSheet.Cells[28, 2] = "D1. Ödənişin təyinatı və əsas / Payment details: Əmək haqqı";
            oSheet.Cells[31, 2] = "D2. Ödənişlə əlaqədar əlavə informasiya / Narrative:";
            oSheet.Cells[35, 2] = "D3. Büdcə təsnifatının kodu:";
            oSheet.Cells[35, 8] = "D4. Büdcə səviyyəsinin kodu:";

            MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["HARDAN"].Value + "'");
            MyData.dtmainrekvizitler=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

            oSheet.Cells[17, 2] = "Adı / Name: “AGLizinq” QSC";
            oSheet.Cells[18, 2] = "Hesab № / Acc. №: " + MyData.dtmainrekvizitler.Rows[0]["a2"].ToString();
            oSheet.Cells[20, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a3"].ToString();

            oSheet.Cells[8, 2] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0]["a4"].ToString();
            oSheet.Cells[9, 2] = "Kodu / Code: " + MyData.dtmainrekvizitler.Rows[0]["a5"].ToString();
            oSheet.Cells[10, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a6"].ToString();
            oSheet.Cells[11, 2] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainrekvizitler.Rows[0]["a7"].ToString();
            oSheet.Cells[13, 2] = "S. W. I. F. T. Bik: " + MyData.dtmainrekvizitler.Rows[0]["a8"].ToString();
          
            oSheet.Cells[17, 8] = "Adı / Name: “AGLizinq” QSC";
            oSheet.Cells[18, 8] = "Hesab № / Acc. №: AZ30AZEG" + Environment.NewLine + "45013944017950107055";
            oSheet.Cells[20, 8] = "VÖEN / Tax İD: 1300616961";

            oSheet.Cells[8, 8] = "Adı / Name: AGBANK ASC";
            oSheet.Cells[9, 8] = "Kodu / Code: 505817";
            oSheet.Cells[10, 8] = "VÖEN / Tax İD: 9900019651";
            oSheet.Cells[11, 8] = "Müxbir hesab / Corr. Acc.: AZ75NABZ" + Environment.NewLine + "01350100000000017944"; ;
            oSheet.Cells[13, 8] = "S. W. I. F. T. Bik: AZEGAZ22";
            
        }

        private void gəlirVergisiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try { File.Copy("Plateshka.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\14%.xlsx", true); }
            catch { MessageBox.Show("'Plateshka.xlsx' tapılmadı."); }

            DateTime dt = Convert.ToDateTime(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["TARİX"].Value);
            string a = MyChange.TarixSozle(dt);
           
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\14%.xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

            oSheet.Cells[2, 2] = "ÖDƏNİŞ TAPŞIRIĞI № " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["No"].Value;
            oSheet.Cells[4, 2] = dt.Day + " " + a + " " + dt.Year + " - ci il";
            oSheet.Cells[22, 8] = "AZN";
            oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["MƏBLƏĞ"].Value;
            oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: ";


            try
            {
                oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: " + MyChange.ReqemToMetn(Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["MƏBLƏĞ"].Value));
            }
            catch { }

            oSheet.Cells[28, 2] = "D1. Ödənişin təyinatı və əsas / Payment details: " + Environment.NewLine + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["TƏYİNAT"].Value;
            oSheet.Cells[31, 2] = "D2. Ödənişlə əlaqədar əlavə informasiya / Narrative:";
            oSheet.Cells[35, 2] = "D3. Büdcə təsnifatının kodu: 111111";
            oSheet.Cells[35, 8] = "D4. Büdcə səviyyəsinin kodu: 1";

            MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["HARDAN"].Value + "'");
            MyData.dtmainrekvizitler=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

            oSheet.Cells[17, 2] = "Adı / Name: “AGLizinq” QSC";
            oSheet.Cells[18, 2] = "Hesab № / Acc. №: " + MyData.dtmainrekvizitler.Rows[0]["a2"].ToString();
            oSheet.Cells[20, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a3"].ToString();

            oSheet.Cells[8, 2] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0]["a4"].ToString();
            oSheet.Cells[9, 2] = "Kodu / Code: " + MyData.dtmainrekvizitler.Rows[0]["a5"].ToString();
            oSheet.Cells[10, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a6"].ToString();
            oSheet.Cells[11, 2] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainrekvizitler.Rows[0]["a7"].ToString();
            oSheet.Cells[13, 2] = "S. W. I. F. T. Bik: " + MyData.dtmainrekvizitler.Rows[0]["a8"].ToString();
            
            oSheet.Cells[17, 8] = "Adı / Name: Bakı şəhəri Lokal gəlirlər Departamenti";
            oSheet.Cells[18, 8] = "Hesab № / Acc. №: AZ17CTRE" + Environment.NewLine + "00000000000002117131";
            oSheet.Cells[20, 8] = "VÖEN / Tax İD: 1403006271";

            oSheet.Cells[8, 8] = "Adı / Name: DXA";
            oSheet.Cells[9, 8] = "Kodu / Code: 210005";
            oSheet.Cells[10, 8] = "VÖEN / Tax İD: 1401555071";
            oSheet.Cells[11, 8] = "Müxbir hesab / Corr. Acc.: AZ41NABZ" + Environment.NewLine + "01360100000000003944"; ;
            oSheet.Cells[13, 8] = "S. W. I. F. T. Bik: CTREAZ22";

        }

        private void txtbankvoen_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a6 Like " + "'%" + txtbankvoen.Text + "%'");
                MyData.dtmainrekvizitler=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);
                try
                {
                    txtbankadi.Text = MyData.dtmainrekvizitler.Rows[0][4].ToString();
                    txtbankkodu.Text = MyData.dtmainrekvizitler.Rows[0][5].ToString();
                    txtbankvoen.Text = MyData.dtmainrekvizitler.Rows[0][6].ToString();
                    txtmuxhesab1.Text = MyData.dtmainrekvizitler.Rows[0][7].ToString().Substring(0, 8);
                    txtmuxhesab2.Text = MyData.dtmainrekvizitler.Rows[0][7].ToString().Substring(9, MyData.dtmainrekvizitler.Rows[0][7].ToString().Length - 9);
                    txtsvift.Text = MyData.dtmainrekvizitler.Rows[0][8].ToString();
                }
                catch { MessageBox.Show("Tapılmadı"); };

                button22.Visible = true;
            }
        }

        private void txtmusterivoen_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a3 Like " + "'%" + txtmusterivoen.Text + "%'");
                MyData.dtmainrekvizitler=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

                try
                {
                    txtmusteriadi.Text = MyData.dtmainrekvizitler.Rows[0][1].ToString();
                    txthesab1.Text = MyData.dtmainrekvizitler.Rows[0][2].ToString().Substring(0, 8);
                    txthesab2.Text = MyData.dtmainrekvizitler.Rows[0][2].ToString().Substring(9, MyData.dtmainrekvizitler.Rows[0][2].ToString().Length - 9); ;
                    txtmusterivoen.Text = MyData.dtmainrekvizitler.Rows[0][3].ToString();
                    txtsexsiyyet.Text = MyData.dtmainrekvizitler.Rows[0][9].ToString();

                    txtbankadi.Text = MyData.dtmainrekvizitler.Rows[0][4].ToString();
                    txtbankkodu.Text = MyData.dtmainrekvizitler.Rows[0][5].ToString();
                    txtbankvoen.Text = MyData.dtmainrekvizitler.Rows[0][6].ToString();
                    txtmuxhesab1.Text = MyData.dtmainrekvizitler.Rows[0][7].ToString().Substring(0, 8);
                    txtmuxhesab2.Text = MyData.dtmainrekvizitler.Rows[0][7].ToString().Substring(9, MyData.dtmainrekvizitler.Rows[0][2].ToString().Length - 9); ;
                    txtsvift.Text = MyData.dtmainrekvizitler.Rows[0][8].ToString();
                }
                catch { MessageBox.Show("Tapılmadı"); };

                MyData.selectCommand("baza.accdb", "Select * From plateshkateyinat WHERE a1 Like " + "'%" + txtmusteriadi.Text + "%'");
                MyData.dtmainTeyinat=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainTeyinat);

                try
                {
                    txtesas.Text = MyData.dtmainTeyinat.Rows[0][1].ToString();

                }
                catch { };

                button22.Visible = true;
            }
        }

        private void cbedv1_SelectedIndexChanged(object sender, EventArgs e)
        {
            button6.Enabled = true;
            button8.Enabled = true; 
            
            MyData.selectCommand("baza.accdb", "Select * From PlateshkaEDV WHERE a1 Like " + "'%" + cbedv1.Text + "%'");
            MyData.dtmainedv=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainedv);

            lbbankadi3.Text = "Bank - " + MyData.dtmainedv.Rows[0]["a4"].ToString() + Environment.NewLine + "H/h - " + MyData.dtmainedv.Rows[0]["a2"].ToString() + Environment.NewLine + "M/h - " + MyData.dtmainedv.Rows[0]["a7"].ToString();
        }

        private void editToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.EditMode = DataGridViewEditMode.EditOnEnter;
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView1.EditMode = DataGridViewEditMode.EditProgrammatically;

            try
            {
                MyData.updateCommand("baza.accdb", "UPDATE plateshkaARXIV SET "
                                                                                     + "TARİX ='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["TARİX"].Value.ToString() + "',"
                                                                                     + "№ ='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["No"].Value.ToString() + "',"
                                                                                     + "HARA ='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["HARA"].Value.ToString() + "',"
                                                                                     + "[ÖDƏNİŞİN TƏYİNATI] ='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["TƏYİNAT"].Value.ToString() + "',"
                                                                                     + "MƏBLƏĞ ='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["MƏBLƏĞ"].Value.ToString() + "',"
                                                                                     + "VALYUTA ='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["VALYUTA"].Value.ToString() + "',"
                                                                                     + "HARDAN ='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["HARDAN"].Value.ToString() + "'"
                                                                                     + " WHERE Код Like '" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Код"].Value.ToString() + "'");

                
                
            }
            catch { MessageBox.Show("Əməliyyat baş tutmadı."); }
        }

        private void txtmanat_TextChanged(object sender, EventArgs e)
        {
            reqemler();
        }

        private void txtedvmanat_TextChanged(object sender, EventArgs e)
        {
            reqemler2();
        }

        private void txtkonvertmanat_TextChanged(object sender, EventArgs e)
        {
            reqemler3();
        }

        private void txtmanatEmekhaqqi_TextChanged(object sender, EventArgs e)
        {
            reqemler4();
        }

        private void txtmanatEmekhaqqi3_TextChanged(object sender, EventArgs e)
        {
            reqemler5();
        }

        private void txtmanatEmekhaqqi5_TextChanged(object sender, EventArgs e)
        {
            reqemler6();
        }

        private void txtmanatEmekhaqqi7_TextChanged(object sender, EventArgs e)
        {
            reqemler7();
        }

        private void dataGridView1_MouseClick(object sender, MouseEventArgs e)
        {
            try
            {
                if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["TƏYİNAT"].Value.ToString() == "Əmək haqqı") { əməkHaqqıToolStripMenuItem.BackColor = Color.Green; siğortaEdən05ToolStripMenuItem.BackColor = Color.Red; siğortaOlunan05ToolStripMenuItem.BackColor = Color.Red; toolStripMenuItem2.BackColor = Color.Red; toolStripMenuItem3.BackColor = Color.Red; gəlirVergisiToolStripMenuItem.BackColor = Color.Red; əDZVToolStripMenuItem.BackColor = Color.Red; ödənişTapşırığıToolStripMenuItem.BackColor = Color.Red; return; } else { əməkHaqqıToolStripMenuItem.BackColor = Color.Red; toolStripMenuItem2.BackColor = Color.Red; toolStripMenuItem3.BackColor = Color.Red; gəlirVergisiToolStripMenuItem.BackColor = Color.Red; siğortaEdən05ToolStripMenuItem.BackColor = Color.Red; siğortaOlunan05ToolStripMenuItem.BackColor = Color.Red; əDZVToolStripMenuItem.BackColor = Color.Green; ödənişTapşırığıToolStripMenuItem.BackColor = Color.Green; }

                if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["TƏYİNAT"].Value.ToString() == "DSMF-na ayırmalar (15%)") { toolStripMenuItem2.BackColor = Color.Green; siğortaEdən05ToolStripMenuItem.BackColor = Color.Red; siğortaOlunan05ToolStripMenuItem.BackColor = Color.Red; əməkHaqqıToolStripMenuItem.BackColor = Color.Red; toolStripMenuItem3.BackColor = Color.Red; gəlirVergisiToolStripMenuItem.BackColor = Color.Red; əDZVToolStripMenuItem.BackColor = Color.Red; ödənişTapşırığıToolStripMenuItem.BackColor = Color.Red; return; } else { əməkHaqqıToolStripMenuItem.BackColor = Color.Red; toolStripMenuItem2.BackColor = Color.Red; toolStripMenuItem3.BackColor = Color.Red; gəlirVergisiToolStripMenuItem.BackColor = Color.Red; siğortaEdən05ToolStripMenuItem.BackColor = Color.Red; siğortaOlunan05ToolStripMenuItem.BackColor = Color.Red; əDZVToolStripMenuItem.BackColor = Color.Green; ödənişTapşırığıToolStripMenuItem.BackColor = Color.Green; }

                if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["TƏYİNAT"].Value.ToString() == "İşəgötürənlərin işsizlikdən siğorta haqqı - (0.5%)") { siğortaEdən05ToolStripMenuItem.BackColor = Color.Green; siğortaOlunan05ToolStripMenuItem.BackColor = Color.Red; toolStripMenuItem3.BackColor = Color.Red; toolStripMenuItem2.BackColor = Color.Red; əməkHaqqıToolStripMenuItem.BackColor = Color.Red; gəlirVergisiToolStripMenuItem.BackColor = Color.Red; əDZVToolStripMenuItem.BackColor = Color.Red; ödənişTapşırığıToolStripMenuItem.BackColor = Color.Red; return; } else { əməkHaqqıToolStripMenuItem.BackColor = Color.Red; toolStripMenuItem2.BackColor = Color.Red; toolStripMenuItem3.BackColor = Color.Red; gəlirVergisiToolStripMenuItem.BackColor = Color.Red; siğortaEdən05ToolStripMenuItem.BackColor = Color.Red; siğortaOlunan05ToolStripMenuItem.BackColor = Color.Red; əDZVToolStripMenuItem.BackColor = Color.Green; ödənişTapşırığıToolStripMenuItem.BackColor = Color.Green; }

                if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["TƏYİNAT"].Value.ToString() == "İşləyənlərin işsizlikdən siğorta haqqı - (0.5%)") { siğortaOlunan05ToolStripMenuItem.BackColor = Color.Green; siğortaEdən05ToolStripMenuItem.BackColor = Color.Red; toolStripMenuItem3.BackColor = Color.Red; toolStripMenuItem2.BackColor = Color.Red; əməkHaqqıToolStripMenuItem.BackColor = Color.Red; gəlirVergisiToolStripMenuItem.BackColor = Color.Red; əDZVToolStripMenuItem.BackColor = Color.Red; ödənişTapşırığıToolStripMenuItem.BackColor = Color.Red; return; } else { əməkHaqqıToolStripMenuItem.BackColor = Color.Red; toolStripMenuItem2.BackColor = Color.Red; toolStripMenuItem3.BackColor = Color.Red; gəlirVergisiToolStripMenuItem.BackColor = Color.Red; siğortaEdən05ToolStripMenuItem.BackColor = Color.Red; siğortaOlunan05ToolStripMenuItem.BackColor = Color.Red; əDZVToolStripMenuItem.BackColor = Color.Green; ödənişTapşırığıToolStripMenuItem.BackColor = Color.Green; }

                if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["TƏYİNAT"].Value.ToString() == "DSMF-na ayırmalar (10%)") { toolStripMenuItem3.BackColor = Color.Green; siğortaEdən05ToolStripMenuItem.BackColor = Color.Red; siğortaOlunan05ToolStripMenuItem.BackColor = Color.Red; toolStripMenuItem2.BackColor = Color.Red; əməkHaqqıToolStripMenuItem.BackColor = Color.Red; gəlirVergisiToolStripMenuItem.BackColor = Color.Red; əDZVToolStripMenuItem.BackColor = Color.Red; ödənişTapşırığıToolStripMenuItem.BackColor = Color.Red; return; } else { əməkHaqqıToolStripMenuItem.BackColor = Color.Red; toolStripMenuItem2.BackColor = Color.Red; toolStripMenuItem3.BackColor = Color.Red; gəlirVergisiToolStripMenuItem.BackColor = Color.Red; siğortaEdən05ToolStripMenuItem.BackColor = Color.Red; siğortaOlunan05ToolStripMenuItem.BackColor = Color.Red; əDZVToolStripMenuItem.BackColor = Color.Green; ödənişTapşırığıToolStripMenuItem.BackColor = Color.Green; }

            } catch { }
            try {əməkHaqqıToolStripMenuItem.BackColor = Color.Red; toolStripMenuItem2.BackColor = Color.Red; toolStripMenuItem3.BackColor = Color.Red; gəlirVergisiToolStripMenuItem.BackColor = Color.Red; siğortaEdən05ToolStripMenuItem.BackColor = Color.Red; siğortaOlunan05ToolStripMenuItem.BackColor = Color.Red; əDZVToolStripMenuItem.BackColor = Color.Green; ödənişTapşırığıToolStripMenuItem.BackColor = Color.Green; } catch { }
        }

        private void toolStripMenuItem11_Click(object sender, EventArgs e)
        {
            try
            {
                MyData.deleteCommand("baza.accdb","DELETE FROM Emekhaqqi WHERE Kod Like '%" + dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["Kod"].Value.ToString() + "%'");
                
                MessageBox.Show("Tapşırıq yerinə yetirildi.");

                emekhaqqirefresh();
            }
            catch { };
           
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            dataGridView2.EditMode = DataGridViewEditMode.EditOnEnter;
        }

        private void dataGridView2_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView2.EditMode = DataGridViewEditMode.EditProgrammatically;
        }

        private void btexcel1_Click(object sender, EventArgs e)
        {
            if (checkBox7.Checked == true) IndividualHesablamaRefresh();
            if (checkBox7.Checked == false) HesablamaRefresh();
            emekhaqqirefresh();
            CemRefresh();
        }

        private void button32_Click(object sender, EventArgs e)
        {
            if (!MyCheck.davamYesNo()) return;

            DateTime dt = dttarixEmekhaqqi.Value.Date;
            string a = MyChange.TarixSozle(dt);

            try { File.Copy("Emek haqqi\\EmekHaqqiCedveli.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Emek Haqqi Cedveli - " + a + ".xlsx", true); }
            catch { MessageBox.Show("'EmekHaqqiCedveli.xlsx' tapılmadı."); }

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Emek Haqqi Cedveli - " + a + ".xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oSheet.Name = a + " " + dt.Year;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();

            oSheet.Cells[8, 2] = dt.Year + "-ci ilin " + a + " ayı üçün hesablanmış əmək haqqı";
            oSheet.Cells[9, 2] = "Cədvəli";
            oSheet.Cells[14, 3] = dataGridView2.Rows[0].Cells["c1"].Value.ToString();
            oSheet.Cells[15, 3] = dataGridView2.Rows[1].Cells["c1"].Value.ToString();
            oSheet.Cells[14, 4] = "2WN14ZK";
            oSheet.Cells[15, 4] = "1F1DQUS";
            oSheet.Cells[16, 4] = "1DQDAGK";
            oSheet.Cells[14, 5] = "Baş menecer";
            oSheet.Cells[15, 5] = "Baş mütəxəssis";
            oSheet.Cells[16, 5] = "Sürücü";
            oSheet.Cells[14, 6] = dataGridView2.Rows[0].Cells["c2"].Value.ToString();
            oSheet.Cells[15, 6] = dataGridView2.Rows[1].Cells["c2"].Value.ToString();
            oSheet.Cells[14, 7] = dataGridView2.Rows[0].Cells["c11"].Value.ToString();
            oSheet.Cells[15, 7] = dataGridView2.Rows[1].Cells["c11"].Value.ToString();
            oSheet.Cells[14, 8] = dataGridView2.Rows[0].Cells["c2"].Value.ToString();
            oSheet.Cells[15, 8] = dataGridView2.Rows[1].Cells["c2"].Value.ToString();
            oSheet.Cells[14, 9] = dataGridView2.Rows[0].Cells["c3"].Value.ToString();
            oSheet.Cells[15, 9] = dataGridView2.Rows[1].Cells["c3"].Value.ToString();
            oSheet.Cells[14, 10] = dataGridView2.Rows[0].Cells["c9"].Value.ToString();
            oSheet.Cells[15, 10] = dataGridView2.Rows[1].Cells["c9"].Value.ToString();
            oSheet.Cells[14, 12] = dataGridView2.Rows[0].Cells["c5"].Value.ToString();
            oSheet.Cells[15, 12] = dataGridView2.Rows[1].Cells["c5"].Value.ToString();
            oSheet.Cells[14, 13] = dataGridView2.Rows[0].Cells["c6"].Value.ToString();
            oSheet.Cells[15, 13] = dataGridView2.Rows[1].Cells["c6"].Value.ToString();
            oSheet.Cells[14, 14] = dataGridView2.Rows[0].Cells["c8"].Value.ToString();
            oSheet.Cells[15, 14] = dataGridView2.Rows[1].Cells["c8"].Value.ToString();
            oSheet.Cells[14, 15] = dataGridView2.Rows[0].Cells["c88"].Value.ToString();
            oSheet.Cells[15, 15] = dataGridView2.Rows[1].Cells["c88"].Value.ToString();
            oSheet.Cells[20, 15] = btCem.Text;
            oSheet.Cells[24, 15] = btPensiya22Faiz2.Text;
        }

        private void txtmanatSigEden05faiz_TextChanged(object sender, EventArgs e)
        {
            reqemler8();
        }

        private void txtmanatSigolunan05faiz_TextChanged(object sender, EventArgs e)
        {
            reqemler9();
        }

        private void siğortaEdən05ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try { File.Copy("Plateshka.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Sigorta Eden 0.5%.xlsx", true); }
            catch { MessageBox.Show("'Plateshka.xlsx' tapılmadı."); }

            DateTime dt = Convert.ToDateTime(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["TARİX"].Value);
            string a = MyChange.TarixSozle(dt);
         
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Sigorta Eden 0.5%.xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

            oSheet.Cells[2, 2] = "ÖDƏNİŞ TAPŞIRIĞI № " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["No"].Value;
            oSheet.Cells[4, 2] = dt.Day + " " + a + " " + dt.Year + "-ci il";
            oSheet.Cells[22, 8] = "AZN";
            oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["MƏBLƏĞ"].Value;
            oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: ";

            try
            {
                oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: " + MyChange.ReqemToMetn(Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["MƏBLƏĞ"].Value));
            }
            catch { }

            oSheet.Cells[28, 2] = "D1. Ödənişin təyinatı və əsas / Payment details: İşəgötürənlərin işsizlikdən siğorta haqqı - (0.5%)";
            oSheet.Cells[31, 2] = "D2. Ödənişlə əlaqədar əlavə informasiya / Narrative:";
            oSheet.Cells[35, 2] = "D3. Büdcə təsnifatının kodu: 123100";
            oSheet.Cells[35, 8] = "D4. Büdcə səviyyəsinin kodu: 4";

            MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["HARDAN"].Value + "'");
            MyData.dtmainrekvizitler=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

            oSheet.Cells[17, 2] = "Adı / Name: “AGLizinq” QSC";
            oSheet.Cells[18, 2] = "Hesab № / Acc. №: " + MyData.dtmainrekvizitler.Rows[0]["a2"].ToString();
            oSheet.Cells[20, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a3"].ToString();

            oSheet.Cells[8, 2] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0]["a4"].ToString();
            oSheet.Cells[9, 2] = "Kodu / Code: " + MyData.dtmainrekvizitler.Rows[0]["a5"].ToString();
            oSheet.Cells[10, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a6"].ToString();
            oSheet.Cells[11, 2] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainrekvizitler.Rows[0]["a7"].ToString();
            oSheet.Cells[13, 2] = "S. W. I. F. T. Bik: " + MyData.dtmainrekvizitler.Rows[0]["a8"].ToString();

            ///////-------------------------Benefisiar alan bank ve musterinin rekvizitleri ucun
            if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["HARA"].Value.ToString() == "BHŞİD")
            {
                oSheet.Cells[17, 8] = "Adı / Name: BHŞİD";
                oSheet.Cells[18, 8] = "Hesab № / Acc. №: AZ42CTRE" + Environment.NewLine + "00000000000007018576";
                oSheet.Cells[20, 8] = "VÖEN / Tax İD: 1600351421";

                oSheet.Cells[8, 8] = "Adı / Name: DXA";
                oSheet.Cells[9, 8] = "Kodu / Code: 210005";
                oSheet.Cells[10, 8] = "VÖEN / Tax İD: 1401555071";
                oSheet.Cells[11, 8] = "Müxbir hesab / Corr. Acc.: AZ41NABZ" + Environment.NewLine + "01360100000000003944"; ;
                oSheet.Cells[13, 8] = "S. W. I. F. T. Bik: CTREAZ22";
            }
            else
            {
                oSheet.Cells[17, 8] = "Adı / Name: DSMF-nin Mərkəzi Sosial Siğorta və Fərdi Uçot İdarəsi";
                oSheet.Cells[18, 8] = "Hesab № / Acc. №: AZ53CTRE" + Environment.NewLine + "00000000000007018572";
                oSheet.Cells[20, 8] = "VÖEN / Tax İD: 1300115511";

                oSheet.Cells[8, 8] = "Adı / Name: DXA";
                oSheet.Cells[9, 8] = "Kodu / Code: 210005";
                oSheet.Cells[10, 8] = "VÖEN / Tax İD: 1401555071";
                oSheet.Cells[11, 8] = "Müxbir hesab / Corr. Acc.: AZ41NABZ" + Environment.NewLine + "01360100000000003944"; ;
                oSheet.Cells[13, 8] = "S. W. I. F. T. Bik: CTREAZ22";
            }

           
        }

        private void siğortaOlunan05ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try { File.Copy("Plateshka.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Sigorta Olunan 0.5%.xlsx", true); }
            catch { MessageBox.Show("'Plateshka.xlsx' tapılmadı."); }

            
            DateTime dt = Convert.ToDateTime(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["TARİX"].Value);
            string a = MyChange.TarixSozle(dt);

            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Sigorta Olunan 0.5%.xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

            oSheet.Cells[2, 2] = "ÖDƏNİŞ TAPŞIRIĞI № " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["No"].Value;
            oSheet.Cells[4, 2] = dt.Day + " " + a + " " + dt.Year + "-ci il";
            oSheet.Cells[22, 8] = "AZN";
            oSheet.Cells[25, 2] = "Məbləğ rəqəmlə: " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["MƏBLƏĞ"].Value;
            oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: ";

            try
            {
                oSheet.Cells[26, 2] = "Məbləğ yazı ilə / İn words: " + MyChange.ReqemToMetn(Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["MƏBLƏĞ"].Value));
            }
            catch { }

            oSheet.Cells[28, 2] = "D1. Ödənişin təyinatı və əsas / Payment details: İşləyənlərin işsizlikdən siğorta haqqı - (0.5%)";
            oSheet.Cells[31, 2] = "D2. Ödənişlə əlaqədar əlavə informasiya / Narrative:";
            oSheet.Cells[35, 2] = "D3. Büdcə təsnifatının kodu: 123200";
            oSheet.Cells[35, 8] = "D4. Büdcə səviyyəsinin kodu: 4";

            MyData.selectCommand("baza.accdb", "Select * From rekvizitler WHERE a1='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["HARDAN"].Value + "'");
            MyData.dtmainrekvizitler=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainrekvizitler);

            oSheet.Cells[17, 2] = "Adı / Name: “AGLizinq” QSC";
            oSheet.Cells[18, 2] = "Hesab № / Acc. №: " + MyData.dtmainrekvizitler.Rows[0]["a2"].ToString();
            oSheet.Cells[20, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a3"].ToString();

            oSheet.Cells[8, 2] = "Adı / Name: " + MyData.dtmainrekvizitler.Rows[0]["a4"].ToString();
            oSheet.Cells[9, 2] = "Kodu / Code: " + MyData.dtmainrekvizitler.Rows[0]["a5"].ToString();
            oSheet.Cells[10, 2] = "VÖEN / Tax İD: " + MyData.dtmainrekvizitler.Rows[0]["a6"].ToString();
            oSheet.Cells[11, 2] = "Müxbir hesab / Corr. Acc.: " + MyData.dtmainrekvizitler.Rows[0]["a7"].ToString();
            oSheet.Cells[13, 2] = "S. W. I. F. T. Bik: " + MyData.dtmainrekvizitler.Rows[0]["a8"].ToString();

            if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["HARA"].Value.ToString() == "BHŞİD")
            {
                oSheet.Cells[17, 8] = "Adı / Name: BHŞİD";
                oSheet.Cells[18, 8] = "Hesab № / Acc. №: AZ42CTRE" + Environment.NewLine + "00000000000007018576";
                oSheet.Cells[20, 8] = "VÖEN / Tax İD: 1600351421";

                oSheet.Cells[8, 8] = "Adı / Name: DXA";
                oSheet.Cells[9, 8] = "Kodu / Code: 210005";
                oSheet.Cells[10, 8] = "VÖEN / Tax İD: 1401555071";
                oSheet.Cells[11, 8] = "Müxbir hesab / Corr. Acc.: AZ41NABZ" + Environment.NewLine + "01360100000000003944"; ;
                oSheet.Cells[13, 8] = "S. W. I. F. T. Bik: CTREAZ22";
            }
            else
            {
                oSheet.Cells[17, 8] = "Adı / Name: DSMF-nin Mərkəzi Sosial Siğorta və Fərdi Uçot İdarəsi";
                oSheet.Cells[18, 8] = "Hesab № / Acc. №: AZ53CTRE" + Environment.NewLine + "00000000000007018572";
                oSheet.Cells[20, 8] = "VÖEN / Tax İD: 1300115511";

                oSheet.Cells[8, 8] = "Adı / Name: DXA";
                oSheet.Cells[9, 8] = "Kodu / Code: 210005";
                oSheet.Cells[10, 8] = "VÖEN / Tax İD: 1401555071";
                oSheet.Cells[11, 8] = "Müxbir hesab / Corr. Acc.: AZ41NABZ" + Environment.NewLine + "01360100000000003944"; ;
                oSheet.Cells[13, 8] = "S. W. I. F. T. Bik: CTREAZ22";
            }

        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) textBox6.Text = Math.Round((Convert.ToDouble(textBox2.Text) / 17050 * 10000), 4).ToString();
        }

        private void textBox6_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) textBox2.Text = Math.Round((Convert.ToDouble(textBox6.Text) * 17000 / 10000), 4).ToString();
        }

        private void txtedv1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                MyData.selectCommand("baza.accdb", "Select * From PlateshkaEDV WHERE a1 Like " + "'%" + txtedv1.Text + "%'");
                MyData.dtmainedv=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainedv);

                try
                {
                    txtedv1.Text = MyData.dtmainedv.Rows[0]["a1"].ToString();
                    txtedv2.Text = MyData.dtmainedv.Rows[0]["a2"].ToString().Substring(0, 8);
                    txtedv22.Text = MyData.dtmainedv.Rows[0]["a2"].ToString().Substring(9, MyData.dtmainedv.Rows[0]["a2"].ToString().Length - 9); ;
                    txtedv3.Text = MyData.dtmainedv.Rows[0]["a3"].ToString();
                    txtedv4.Text = MyData.dtmainedv.Rows[0]["a4"].ToString();
                    txtedv5.Text = MyData.dtmainedv.Rows[0]["a5"].ToString();
                    txtedv6.Text = MyData.dtmainedv.Rows[0]["a6"].ToString();
                    txtedv7.Text = MyData.dtmainedv.Rows[0]["a7"].ToString().Substring(0, 8);
                    txtedv77.Text = MyData.dtmainedv.Rows[0]["a7"].ToString().Substring(9, MyData.dtmainedv.Rows[0]["a7"].ToString().Length - 9); ; ;
                    txtedv8.Text = MyData.dtmainedv.Rows[0]["a8"].ToString();
                    txtTesnifatKod.Text = MyData.dtmainedv.Rows[0]["a9"].ToString();
                    txtSeviyyeKod.Text = MyData.dtmainedv.Rows[0]["a10"].ToString();
                }
                catch { MessageBox.Show("Tapılmadı"); return; };

                MyData.selectCommand("baza.accdb", "Select * From plateshkateyinat where a1 Like '%" + txtedv1.Text + "%'");
                MyData.dtmainTeyinat=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainTeyinat);
                try { txtteyinatEDV.Text = MyData.dtmainTeyinat.Rows[0]["a2"].ToString(); }
                catch { }
            }
        }

        private void button21_Click(object sender, EventArgs e)
        {
            //comboBox2.Text = "";
            //cbedv2.Text = "";
            //txtHara.Text = "";
            //txtteyinat.Text = "";
            //txtMebleg.Text = "";
            //dtBaslama.Value = DateTime.Now.AddYears(-1);
            //dtBitme.Value = DateTime.Now;
            //txtkonvertmanat.Text = "";

            emekhaqqirefresh();
            ARXIV();
            plateshkanomreRefresh();
            EDVRefresh();
            EDVnomreRefresh();
            rekvizitlerRefresh();
            CemRefresh();
            anonsrefresh();
        }

        private void cbSatilanValyutaHesab_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbSatilanValyutaHesab.Text.Substring(cbAlinanValyutaHesab.Text.Length - 3, 3) != "AZN") lbBasliq.Text = "“AGBANK”  AÇIQ SƏHMDAR  CƏMİYYƏTİ" + Environment.NewLine + "xarici valyutanın Satılması üçün" + Environment.NewLine + "SİFARİŞ";
            else lbBasliq.Text = "“AGBANK”  AÇIQ SƏHMDAR  CƏMİYYƏTİ" + Environment.NewLine + "xarici valyutanın Alınması üçün" + Environment.NewLine + "SİFARİŞ";
        }

        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {
            string k = txtmanat.Text;
            
            cbedv2.Text = comboBox2.Text;
            txtmanat.Text = Math.Round((Convert.ToDouble(k) / 1.18), 2, MidpointRounding.AwayFromZero).ToString();
            txtedvmanat.Text = Math.Round((Convert.ToDouble(k) - Convert.ToDouble(k) / 1.18), 2, MidpointRounding.AwayFromZero).ToString();
            txtedvteyinat.Text = txtteyinat.Text;
            dtedvtarix.Text = dttarix.Text;
        }

        private void toolStripMenuItem5_Click(object sender, EventArgs e)
        {
            string k = txtedvmanat.Text;
            txtedvmanat.Text = Math.Round((Convert.ToDouble(k) - Convert.ToDouble(k) / 1.18), 2, MidpointRounding.AwayFromZero).ToString();
        }

        private void txtteyinat_TextChanged(object sender, EventArgs e)
        {
            txtedvteyinat.Text = txtteyinat.Text;
        }

        private void button24_Click(object sender, EventArgs e)
        {

            try { File.Copy("New Emphty.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Rekvizitlər.doc", true); }
            catch { MessageBox.Show("'\\192.168.10.5\\Common\\AGLizinq\\New Emphty.doc' tapılmadı."); }

            object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\Rekvizitlər.doc";
            Microsoft.Office.Interop.Word._Application oWord = new Microsoft.Office.Interop.Word.Application();
            object oMissing = Type.Missing;
            oWord.Visible = true;
            oWord.Documents.Open(FileName);
            oWord.Selection.TypeText(txtmusteriadi.Text + " - nin " + Environment.NewLine + "Bank rekvizitləri" + Environment.NewLine + Environment.NewLine + Environment.NewLine);
            oWord.Selection.TypeText("ADI: " + txtmusteriadi.Text + Environment.NewLine + "VÖEN: " + txtmusterivoen.Text + Environment.NewLine + "H/h: " + txthesab1.Text + txthesab2.Text + Environment.NewLine + "Ş/V: " + txtsexsiyyet.Text + Environment.NewLine + Environment.NewLine);
            oWord.Selection.TypeText("BANK Adı: " + txtbankadi.Text + Environment.NewLine + "VÖEN: " + txtbankvoen.Text + Environment.NewLine + "KODU: " + txtbankkodu.Text + Environment.NewLine + "M/h: " + txtmuxhesab1.Text + txtmuxhesab2.Text + Environment.NewLine + "S.W.I.F.T. Bik: " + txtsvift.Text);
            //oWord.PrintOut();
            oWord.ActiveDocument.Save();
            //oWord.Quit();
        }

        private void button28_Click(object sender, EventArgs e)
        {
            if (!MyCheck.davamYesNo()) return;

            DateTime dt = dttarixEmekhaqqi.Value.Date;
            string a = MyChange.TarixSozle(dt);
           
            try { File.Copy("Emek haqqi\\Tabel.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Tabel - " + a + " " + dt.Year + ".xlsx", true); }
            catch { MessageBox.Show("'Tabel.xlsx' tapılmadı."); }

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Tabel - " + a + " " + dt.Year + ".xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            oSheet.Name = a + " " + dt.Year;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();

            oSheet.Cells[6, 1] = "'" + a.ToUpper() +  " " + dt.Year;
            oSheet.Cells[8, 1] = "VÖEN 1300616961";
            oSheet.Cells[22, 6] = dataGridView2.Rows[0].Cells["c4"].Value.ToString();
            oSheet.Cells[24, 6] = dataGridView2.Rows[1].Cells["c4"].Value.ToString();

            //oSheet.Columns.AutoFit();
            //oSheet.Rows.AutoFit();
        }

        private void button29_Click(object sender, EventArgs e)
        {
            if (txtedv1.Text == "") { MessageBox.Show("Müştərinin Adı qeyd olunmayıb"); return; }
            if (txtedv2.Text == "") { MessageBox.Show("Müştərinin Hesabı qeyd olunmayıb"); return; }
            if (txtedv22.Text == "") { MessageBox.Show("Müştərinin Hesabı qeyd olunmayıb"); return; }
            if (txtedv3.Text == "") { MessageBox.Show("Müştərinin VÖEN-i qeyd olunmayıb"); return; }
            if (txtedv4.Text == "") { MessageBox.Show("Bankın Adı qeyd olunmayıb"); return; }
            if (txtedv5.Text == "") { MessageBox.Show("Bankın Kodu qeyd olunmayıb"); return; }
            if (txtedv6.Text == "") { MessageBox.Show("Bankın VÖEN-i qeyd olunmayıb"); return; }
            if (txtedv7.Text == "") { MessageBox.Show("Bankın Müxbir Hesabı qeyd olunmayıb"); return; }
            if (txtedv77.Text == "") { MessageBox.Show("Bankın Müxbir Hesabı qeyd olunmayıb"); return; }
            if (txtedv8.Text == "") { MessageBox.Show("Bankın S.W.I.F.T-i qeyd olunmayıb"); return; }

            try
            {
                MyData.updateCommand("baza.accdb", "UPDATE PlateshkaEDV SET "
                                                                                     + "a1 ='" + txtedv1.Text + "',"
                                                                                     + "a2 ='" + txtedv2.Text + " " + txtedv22.Text + "',"
                                                                                     + "a3 ='" + txtedv3.Text + "',"
                                                                                     + "a4 ='" + txtedv4.Text + "',"
                                                                                     + "a5 ='" + txtedv5.Text + "',"
                                                                                     + "a6 ='" + txtedv6.Text + "',"
                                                                                     + "a7 ='" + txtedv7.Text + " " + txtedv77.Text + "',"
                                                                                     + "a8 ='" + txtedv8.Text + "',"
                                                                                     + "a9 ='" + txtTesnifatKod.Text + "',"
                                                                                     + "a10 ='" + txtSeviyyeKod.Text + "'"
                                                                                     + " WHERE a1 Like '" + txtedv1.Text + "'");

                MyData.updateCommand("baza.accdb", "UPDATE plateshkateyinat SET a2='" + txtteyinatEDV.Text + "' WHERE a1 Like '%" + txtedv1.Text + "%'");
                
                MessageBox.Show("Yadda saxlanıldı.");
            }
            catch { MessageBox.Show("Əməliyyatda səhv var." + Environment.NewLine + " - Bu adda müştəri yaddaşda yoxdur"); }
        }

        private void dttarixEmekhaqqi_ValueChanged(object sender, EventArgs e)
        {
            DateTime dt = dttarixEmekhaqqi.Value.Date;
            string t = MyChange.TarixSozle(dt);

            textBox8.Text = dt.Year + "-ci il " + t + " ayı üçün gəlir vergisi.";
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            lbAnons.Left -= 2;
            if (lbAnons.Right < 0) lbAnons.Left = base.Width;
        }

        private void txtmanatIsegoturen1faiz_TextChanged(object sender, EventArgs e)
        {
            reqemler10();
        }

        private void txtmanatIsciler1faiz_TextChanged(object sender, EventArgs e)
        {
            reqemler11();
        }

        private void button31_Click(object sender, EventArgs e)
        {
            if (!MyCheck.davamYesNo()) return;

            DateTime dt = dttarixEmekhaqqi.Value.Date;
            string a = MyChange.TarixSozle(dt);
            
            try { File.Copy("Emek haqqi\\Payment.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\" + a + " " + dt.Year + ".xlsx", true); }
            catch { MessageBox.Show("'Payment.xlsx' tapılmadı."); }

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\" + a + " " + dt.Year + ".xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();

            oSheet.Cells[1, 3] = dataGridView2.Rows[0].Cells["c4"].Value.ToString();
            oSheet.Cells[2, 3] = dataGridView2.Rows[1].Cells["c4"].Value.ToString();
            oSheet.Cells[1, 4] = a + " emekhaqqı";
            oSheet.Cells[2, 4] = a + " emekhaqqı";
            oSheet.Cells[1, 5] = "EH";
            oSheet.Cells[2, 5] = "EH";
        }

        private void TtxtTeyinat_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try { ARXIV(); } catch { }
            }
        }

        private void TxtMebleg_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try { ARXIV(); } catch { }
            }
        }

        private void DtBaslama_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try { ARXIV(); } catch { }
            }
        }

        private void DtBitme_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try { ARXIV(); } catch { }
            }
        }

        private void button33_Click(object sender, EventArgs e)
        {
            //try
            //{
                int s = 0, k = 0, a = 0;

                try { File.Copy("Bos.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Odenis Tapsiriqlari.xlsx", true); }
                catch { MessageBox.Show("'Portfel.xlsm' tapılmadı."); }

                //Get a new workbook.
                oXL = new Excel.Application();
                oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Odenis Tapsiriqlari.xlsx"));
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
                oSheet.Cells[1, a].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Cells[1, 1] = "Nö";

                for (k = 0; k < dataGridView1.Rows.Count; k++)
                {
                    oSheet.Cells[k + 2, a] = dataGridView1.Rows[k].Cells[s].Value;
                    if (k != dataGridView1.Rows.Count) 
                    { 
                        oSheet.Cells[k + 2, 1] = k + 1;
                        oSheet.Cells[k + 2, 1].Borders.LineStyle = Excel.Constants.xlSolid;
                    }

                    oSheet.Cells[k + 2, s + 2].Borders.LineStyle = Excel.Constants.xlSolid;
                }
            }
            //}
            //catch { }
        }
    }
}
