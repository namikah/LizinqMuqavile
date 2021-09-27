using Nsoft;
using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;

namespace Lizinq_Muqavile
{
    public partial class Bildiris : Form
    {
        public Bildiris()
        {
            InitializeComponent();
        }

        public void WordBildiris()
        {
            try { File.Copy("Bildiris.doc", "X:\\Umumi Senedler\\••• SƏNƏDLƏR •••\\BILDIRIŞ\\Bildiris Sistem\\" + dttarix.Text + " " + txtlizinqalan.Text + ".doc", true); }
            catch { }

            try
            {
                object FileName = "X:\\Umumi Senedler\\••• SƏNƏDLƏR •••\\BILDIRIŞ\\Bildiris Sistem\\" + dttarix.Text + " " + txtlizinqalan.Text + ".doc";

                Word.Application word = new Word.Application();
                Word.Document doc = null;
                object missing = System.Type.Missing;
                object readOnly = false;
                object isVisible = false;
                word.Visible = true;

                doc = word.Documents.Open(ref FileName);
                doc.Activate();

                DateTime dt = dttarix.Value.Date;
                DateTime dt2 = dttarix2.Value.Date;

                string a = MyChange.TarixSozle(dt2); //muqavile tarixi
                string b = MyChange.TarixSozle(dt); //hazirki tarix

                MyChange.FindAndReplace(word, "0000", dt.Day + " " + b + " " + dt.Year + " - ci il");
                MyChange.FindAndReplace(word, "0000", dt.Day + " " + b + " " + dt.Year + " - ci il");
                MyChange.FindAndReplace(word, "000", txtlayihe.Text);
                MyChange.FindAndReplace(word, "000", txtlayihe.Text);
                MyChange.FindAndReplace(word, "000", txtlayihe.Text);
                MyChange.FindAndReplace(word, "111", dt2.Day + " " + a + " " + dt2.Year + " - ci il");
                MyChange.FindAndReplace(word, "222", txtlizinqalan.Text);
                MyChange.FindAndReplace(word, "222", txtlizinqalan.Text);
                MyChange.FindAndReplace(word, "222", txtlizinqalan.Text);
                MyChange.FindAndReplace(word, "333", comboBox1.Text);
                MyChange.FindAndReplace(word, "333", comboBox1.Text);
                MyChange.FindAndReplace(word, "333", comboBox1.Text);
                MyChange.FindAndReplace(word, "444", cbmuqavileFormasi.Text);
                MyChange.FindAndReplace(word, "555", "Ş/V " + txtsexsiyyet.Text);
                MyChange.FindAndReplace(word, "666", txtunvan.Text);
                MyChange.FindAndReplace(word, "777", txtGecikmisesasborc.Text + " (" + txtgecikmisesasborcHERF.Text + ")");
                MyChange.FindAndReplace(word, "888", txtGecikmisfaizborc.Text + " (" + txtfaizborcHERF.Text + ")");
                MyChange.FindAndReplace(word, "1234", txtcerime.Text + " (" + txtcerimeHERF.Text + ")");
                MyChange.FindAndReplace(word, "1235", txtqaliqesas.Text + " (" + txtqaliqesasHERF.Text + ")");
                MyChange.FindAndReplace(word, "999", txtumumiborc.Text + " (" + txtumumiborcHERF.Text + ")");
                MyChange.FindAndReplace(word, "11111111", txtavadanliq.Text);
                MyChange.FindAndReplace(word, "22222222", txtlizinqmebleg.Text + " (" + txtlizinqmeblegHerf.Text + ")");

            doc.Save();
            }
            catch (Exception ex){ MessageBox.Show(ex.Message + Environment.NewLine + "'X:\\Umumi Senedler\\••• SƏNƏDLƏR •••\\BILDIRIŞ\\Bildiris Sistem\\" + dttarix.Text + " " + txtlizinqalan.Text + ".doc' tapılmadı." + Environment.NewLine + "Lizinq alanın adında naməlum işarələr (/:*?<>) aşkar olundu."); }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            WordBildiris();
        }

        private void txtesasborc_TextChanged(object sender, EventArgs e)
        {
            try
            {
                txtumumiborc.Text = (Convert.ToDouble(txtqaliqesas.Text) + Convert.ToDouble(txtGecikmisesasborc.Text) + Convert.ToDouble(txtGecikmisfaizborc.Text) + Convert.ToDouble(txtcerime.Text)).ToString();
                txtgecikmisesasborcHERF.Text = MyChange.ReqemToMetn(Convert.ToDouble(txtGecikmisesasborc.Text));
            }
            catch { }
        }

        private void txtfaizborc_TextChanged(object sender, EventArgs e)
        {
            try
            {
                txtumumiborc.Text = (Convert.ToDouble(txtqaliqesas.Text) + Convert.ToDouble(txtGecikmisesasborc.Text) + Convert.ToDouble(txtGecikmisfaizborc.Text) + Convert.ToDouble(txtcerime.Text)).ToString();
                txtfaizborcHERF.Text = MyChange.ReqemToMetn(Convert.ToDouble(txtGecikmisfaizborc.Text));
            }
            catch { }
        }

        private void txtumumiborc_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    txtumumiborc.Text = (Convert.ToDouble(txtGecikmisesasborc.Text) + Convert.ToDouble(txtGecikmisfaizborc.Text) + Convert.ToDouble(txtcerime.Text)).ToString();
                }
            }
            catch { }
        }

        private void txtumumiborc_TextChanged(object sender, EventArgs e)
        {
            txtumumiborcHERF.Text = MyChange.ReqemToMetn(Convert.ToDouble(txtumumiborc.Text));
        }

        private void txtlayihe_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                    txtlizinqalan.Text = "";
                    txtsexsiyyet.Text = "AZE № ";
                    txtunvan.Text = "";
                    txtGecikmisesasborc.Text = "0";
                    txtGecikmisfaizborc.Text = "0";
                    txtlizinqmebleg.Text = "0";
                    txtavadanliq.Text = "";
                    txtumumiborc.Text = "0";

                    //Lizinq alan ve Layihenin oxunmasi ucun lizinq alan
                    try
                    {
                    MyData.selectCommand("baza.accdb", "SELECT * FROM Etibarnameneqliyyat WHERE c4 Like '%" + txtlayihe.Text + "%'");
                    MyData.dtmain = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmain);

                        if (MyData.dtmain.Rows.Count > 0)
                        {
                            txtlayihe.Text = MyData.dtmain.Rows[0]["c4"].ToString();
                            txtlizinqalan.Text = MyData.dtmain.Rows[0]["c3"].ToString();
                            txtavadanliq.Text = MyData.dtmain.Rows[0]["c2"].ToString();
                        }
                    }
                    catch { }

                  
                    try
                    {
                            DateTime dt = DateTime.Now;
                            String name = "licschkre";
                            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                            "2.xlsx" +
                                            ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

                            OleDbConnection con = new OleDbConnection(constr);
                            OleDbCommand oconn = new OleDbCommand("Select Adı, [V#K#Qalıq], [V#K#% məbləği], [Dəbbə məbləği], [Cərimə % məbləği], [Verilmə tarixi], Layihe, [Lizinqin məbləği], [% məbləği], Qalıq From [" + name + "$] WHERE Layihe Like '%" + txtlayihe.Text + "%'", con);
                            con.Open();
                            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                            DataTable data = new DataTable();
                            data.Clear();
                            sda.Fill(data);
                            con.Close();

                            dttarix2.Text = data.Rows[0]["Verilmə tarixi"].ToString();
                            txtGecikmisesasborc.Text = data.Rows[0]["V#K#Qalıq"].ToString();
                            txtqaliqesas.Text = data.Rows[0]["Qalıq"].ToString();
                            txtGecikmisfaizborc.Text = Math.Round(Convert.ToDouble(data.Rows[0]["% məbləği"]) + Convert.ToDouble(data.Rows[0]["V#K#% məbləği"]) + (Convert.ToDouble(data.Rows[0]["Cərimə % məbləği"]) + Convert.ToDouble(data.Rows[0]["Dəbbə məbləği"])) / 2, 2).ToString();
                            txtcerime.Text = Math.Round((Convert.ToDouble(data.Rows[0]["Cərimə % məbləği"]) + Convert.ToDouble(data.Rows[0]["Dəbbə məbləği"])) / 2, 2).ToString();
                            dttarix3.Text = dt.AddDays(10).ToShortDateString();
                            txtlizinqmebleg.Text = data.Rows[0]["Lizinqin məbləği"].ToString();

                            if (txtlizinqalan.Text == "") { txtlizinqalan.Text = data.Rows[0]["Adı"].ToString(); txtlayihe.Text = data.Rows[0]["Layihe"].ToString(); }
                       
                    }
                    catch { }

                    //sexsiyyet ve unvanin oxunmasi ucun
                    try
                    {
                            MyData.selectCommand("baza.accdb", "SELECT * FROM etibarnamesurucu WHERE a1 Like '%" + txtlizinqalan.Text + "%'");
                    MyData.dtmain.Clear();
                    MyData.oledbadapter1.Fill(MyData.dtmain);

                            if (MyData.dtmain.Rows.Count > 0)
                            {
                                txtsexsiyyet.Text = MyData.dtmain.Rows[0]["a2"].ToString();
                                txtunvan.Text = MyData.dtmain.Rows[0]["a4"].ToString();
                            }
                    }
                    catch { }

                try
                {
                    if (txtlizinqalan.Text == "" || txtumumiborc.Text == "0") 
                    {
                        String name2 = "Кредитный портфель";
                        String constr2 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                        "1.xlsx" +
                                        ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

                        OleDbConnection con2 = new OleDbConnection(constr2);
                        OleDbCommand oconn2 = new OleDbCommand("Select [Номер контракта], [Наименование клиента], [Дата заключения контракта], [Дата окончания контракта], [Сумма контракта в манатном эквиваленте], [Остаток основного долга на дату в манатном эквиваленте], [Остаток просрочки на дату в манатном эквиваленте], [Начисленные проценты в манатном эквиваленте], [Просроченные проценты в манатном эквиваленте], [Штрафные проценты в манатном эквиваленте], [Срок просрочки процентов], [Последняя дата погашения процентов (или просроченных процентов)] From [" + name2 + "$] WHERE [Номер контракта] LIKE '%" + txtlayihe.Text + "%'", con2);
                        con2.Open();

                        OleDbDataAdapter abc = new OleDbDataAdapter(oconn2);
                        DataTable dataABC = new DataTable();
                        dataABC.Clear();
                        abc.Fill(dataABC);

                        dttarix2.Text = dataABC.Rows[0]["Дата заключения контракта"].ToString();
                        txtGecikmisesasborc.Text = dataABC.Rows[0]["Остаток просрочки на дату в манатном эквиваленте"].ToString();
                        txtqaliqesas.Text = dataABC.Rows[0]["Остаток основного долга на дату в манатном эквиваленте"].ToString();
                        txtGecikmisfaizborc.Text = Math.Round(Convert.ToDouble(dataABC.Rows[0]["Начисленные проценты в манатном эквиваленте"]) + Convert.ToDouble(dataABC.Rows[0]["Просроченные проценты в манатном эквиваленте"]) + (Convert.ToDouble(dataABC.Rows[0]["Штрафные проценты в манатном эквиваленте"])) / 2, 2).ToString();
                        txtcerime.Text = dataABC.Rows[0]["Штрафные проценты в манатном эквиваленте"].ToString();
                        txtlizinqmebleg.Text = dataABC.Rows[0]["Сумма контракта в манатном эквиваленте"].ToString();

                        txtlizinqalan.Text = dataABC.Rows[0]["Наименование клиента"].ToString(); txtlayihe.Text = dataABC.Rows[0]["Номер контракта"].ToString();
                    }
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
                    txtlayihe.Text = "";
                    txtsexsiyyet.Text = "AZE № ";
                    txtunvan.Text = "";
                    txtGecikmisesasborc.Text = "0";
                    txtGecikmisfaizborc.Text = "0";
                    txtlizinqmebleg.Text = "0";
                    txtavadanliq.Text = "";

                    //Lizinq alan ve Layihenin oxunmasi ucun lizinq alan
                    try
                    {
                        DataTable dtMain = new DataTable();
                        MyData.selectCommand("baza", "SELECT * FROM Etibarnameneqliyyat WHERE c3 Like '%" + txtlizinqalan.Text + "%'");
                        MyData.dtmain = new DataTable();
                        MyData.oledbadapter1.Fill(MyData.dtmain);

                        if (MyData.dtmain.Rows.Count > 0)
                        {
                            txtlizinqalan.Text = MyData.dtmain.Rows[0]["c3"].ToString();
                            txtlayihe.Text = MyData.dtmain.Rows[0]["c4"].ToString();
                            txtavadanliq.Text = MyData.dtmain.Rows[0]["c2"].ToString();
                        }
                    }
                    catch { }

                    DateTime dt = DateTime.Now;
                    String name = "licschkre";
                    String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    "2.xlsx" +
                                    ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

                    OleDbConnection con = new OleDbConnection(constr);
                    OleDbCommand oconn = new OleDbCommand("Select Adı, [V#K#Qalıq], [V#K#% məbləği], [Dəbbə məbləği], [Cərimə % məbləği], [Verilmə tarixi], Layihe, [Lizinqin məbləği], [% məbləği], Qalıq From [" + name + "$] WHERE Layihe Like '%" + txtlayihe.Text + "%'", con);
                    con.Open();
                    OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                    DataTable data = new DataTable();
                    data.Clear();
                    sda.Fill(data);
                    con.Close();

                    dttarix2.Text = data.Rows[0]["Verilmə tarixi"].ToString();
                    txtGecikmisesasborc.Text = data.Rows[0]["V#K#Qalıq"].ToString();
                    txtqaliqesas.Text = data.Rows[0]["Qalıq"].ToString();
                    txtGecikmisfaizborc.Text = Math.Round(Convert.ToDouble(data.Rows[0]["% məbləği"]) + Convert.ToDouble(data.Rows[0]["V#K#% məbləği"]) + (Convert.ToDouble(data.Rows[0]["Cərimə % məbləği"]) + Convert.ToDouble(data.Rows[0]["Dəbbə məbləği"])) / 2, 2).ToString();
                    txtcerime.Text = Math.Round((Convert.ToDouble(data.Rows[0]["Cərimə % məbləği"]) + Convert.ToDouble(data.Rows[0]["Dəbbə məbləği"])) / 2, 2).ToString();
                    dttarix3.Text = dt.AddDays(10).ToShortDateString();
                    txtlizinqmebleg.Text = data.Rows[0]["Lizinqin məbləği"].ToString();

                    if (txtlayihe.Text == "") { txtlayihe.Text = data.Rows[0]["Layihe"].ToString(); txtlizinqalan.Text = data.Rows[0]["Adı"].ToString(); }

                }
                catch { MessageBox.Show("- gecikmis faiz borcunda sehv ola biler." + Environment.NewLine + "- cerime borcunda sehv ola biler"); }

                //sexsiyyet ve unvanin oxunmasi ucun
                try
                {
                    DataTable dtMain = new DataTable();
                    MyData.selectCommand("baza", "SELECT * FROM etibarnamesurucu WHERE a1 Like '%" + txtlizinqalan.Text + "%'");
                    MyData.dtmain = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmain);

                    txtsexsiyyet.Text = MyData.dtmain.Rows[0]["a2"].ToString();
                    txtunvan.Text = MyData.dtmain.Rows[0]["a4"].ToString();
                }
                catch { }

            }
        }

        private void txtlizinqmebleg_TextChanged(object sender, EventArgs e)
        {
            try { txtlizinqmeblegHerf.Text = MyChange.ReqemToMetn(Convert.ToDouble(txtlizinqmebleg.Text)); } catch{ }
        }

        private void txtcerime_TextChanged(object sender, EventArgs e)
        {
            try
            {
                txtumumiborc.Text = (Convert.ToDouble(txtqaliqesas.Text) + Convert.ToDouble(txtGecikmisesasborc.Text) + Convert.ToDouble(txtGecikmisfaizborc.Text) + Convert.ToDouble(txtcerime.Text)).ToString();
                txtcerimeHERF.Text = MyChange.ReqemToMetn(Convert.ToDouble(txtcerime.Text));
            }
            catch { }
        }

        private void txtqaliqesas_TextChanged(object sender, EventArgs e)
        {
            try
            {
                txtumumiborc.Text = (Convert.ToDouble(txtqaliqesas.Text) + Convert.ToDouble(txtGecikmisesasborc.Text) + Convert.ToDouble(txtGecikmisfaizborc.Text) + Convert.ToDouble(txtcerime.Text)).ToString();
                txtqaliqesasHERF.Text = MyChange.ReqemToMetn(Convert.ToDouble(txtqaliqesas.Text));
            }
            catch { }
        }

        private void Bildiris_Load(object sender, EventArgs e)
        {

        }
    }
}
