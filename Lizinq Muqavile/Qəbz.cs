using Nsoft;
using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;

namespace Lizinq_Muqavile
{
    public partial class Qəbz : Form
    {
        Excel.Application oXL;
        Excel._Workbook oWB;
        Excel._Worksheet oSheet;

        public Qəbz()
        {
            InitializeComponent();
        }


        public void reqemler()      //------reqem yazi ile---------------------------------------------------------------
        {
            try
            {
                txt2.Text = MyChange.ReqemToMetn(Convert.ToDouble(txt1.Text));
            }
            catch { }
        }

        private void txt1_TextChanged(object sender, EventArgs e)
        {
            reqemler();
        }

        private void txtlizinqalan_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try { txtlizinqalan.Text = txtlizinqalan.Text.ToUpper(MyChange.DilDeyisme); }
                catch { }

                try
                {
                    String name = "licschkre";
                    String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    "2.xlsx" +
                                    ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

                    OleDbConnection con = new OleDbConnection(constr);
                    OleDbCommand oconn = new OleDbCommand("Select Adı, Layihe, [Qrafikda olan məbləğ], [Lizinq hesabı], [V#K#Qalıq], [V#K#% məbləği], [Dəbbə məbləği], [Cərimə % məbləği] From [" + name + "$] WHERE Adı Like" + "'%" + txtlizinqalan.Text + "%'", con);
                    con.Open();
                    OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                    DataTable data = new DataTable();
                    data.Clear();
                    sda.Fill(data);
                    con.Close();

                    string gecikme = (Convert.ToDouble(data.Rows[0]["V#K#Qalıq"]) + Convert.ToDouble(data.Rows[0]["V#K#% məbləği"]) + Convert.ToDouble(data.Rows[0]["Dəbbə məbləği"]) + Convert.ToDouble(data.Rows[0]["Cərimə % məbləği"])).ToString();
                    txt1.Text = gecikme;
                    if (gecikme == "0") txt1.Text = data.Rows[0]["Qrafikda olan məbləğ"].ToString();

                    txtlizinqalan.Text = data.Rows[0]["Adı"].ToString();
                    txtodeyen.Text = data.Rows[0]["Adı"].ToString();
                    txthesabnomresi.Text = "5430" + data.Rows[0]["Lizinq hesabı"].ToString().Substring(5, 9);
                    txtteyinat.Text = (data.Rows[0]["Layihe"].ToString() + " Lizinq ödənişi").Replace("'","");
                    
                }
                catch { }
            }
        }

        private void txtteyinat_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try { txtteyinat.Text = txtteyinat.Text.ToUpper(MyChange.DilDeyisme); }
                catch { }

                try
                {
                    String name = "licschkre";
                    String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    "2.xlsx" +
                                    ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

                    OleDbConnection con = new OleDbConnection(constr);
                    OleDbCommand oconn = new OleDbCommand("Select Adı, Layihe, [Qrafikda olan məbləğ], [Lizinq hesabı], [V#K#Qalıq], [V#K#% məbləği], [Dəbbə məbləği], [Cərimə % məbləği] From [" + name + "$] WHERE Layihe Like " + "'%" + txtteyinat.Text + "%'", con);
                    con.Open();
                    OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                    DataTable data = new DataTable();
                    data.Clear();
                    sda.Fill(data);
                    con.Close();

                    string gecikme = (Convert.ToDouble(data.Rows[0]["V#K#Qalıq"]) + Convert.ToDouble(data.Rows[0]["V#K#% məbləği"]) + Convert.ToDouble(data.Rows[0]["Dəbbə məbləği"]) + Convert.ToDouble(data.Rows[0]["Cərimə % məbləği"])).ToString();
                    txt1.Text = gecikme;
                    if (gecikme == "0") txt1.Text = data.Rows[0]["Qrafikda olan məbləğ"].ToString();

                    txtlizinqalan.Text = data.Rows[0]["Adı"].ToString();
                    txtodeyen.Text = data.Rows[0]["Adı"].ToString();
                    txthesabnomresi.Text = "5430" + data.Rows[0]["Lizinq hesabı"].ToString().Substring(5, 9);
                    txtteyinat.Text = (data.Rows[0]["Layihe"].ToString() + " Lizinq ödənişi").Replace("'","");
                }
                catch { }
            }
        }

        private void txtodeyen_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    txtodeyen.Text = txtodeyen.Text.Substring(0, 1).ToUpper(MyChange.DilDeyisme) + txtodeyen.Text.Substring(1, txtodeyen.Text.Length - 1).ToLower(MyChange.DilDeyisme);
                }
                catch { }

                try
                {
                    MyData.selectCommand("baza.accdb", "Select * from etibarnamesurucu where a1 Like " + "'%" + txtodeyen.Text + "%'");
                    MyData.dtmain.Clear();
                    MyData.oledbadapter1.Fill(MyData.dtmain);

                    txtodeyen.Text = MyData.dtmain.Rows[0]["a1"].ToString();
                }
                catch { }
            }
        }

        private void Btprint_Click(object sender, EventArgs e)
        {

            if (!MyCheck.davamYesNo()) return;

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open(Application.StartupPath + "\\Qebz.xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

            DateTime dt = dttarix.Value.Date;
            string a = MyChange.TarixSozle(dt);

            oSheet.Cells[4, "AJ"] = dt.Day + " " + a + " " + dt.Year + "-ci il";
            oSheet.Cells[5, "AJ"] = txthesabnomresi.Text;
            oSheet.Cells[6, "AJ"] = txt1.Text;
            oSheet.Cells[7, "AJ"] = txt2.Text;
            oSheet.Cells[8, "AJ"] = txtteyinat.Text;
            oSheet.Cells[9, "AJ"] = txtodeyen.Text;
        }
    }
}
