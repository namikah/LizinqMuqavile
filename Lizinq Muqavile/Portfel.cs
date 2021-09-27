using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Xml;
using System.IO;
using System.Globalization;

namespace Lizinq_Muqavile
{
    public partial class Portfel : Form
    {
        public Portfel()
        {
            InitializeComponent();
        }

        private void Portfel_Load(object sender, EventArgs e)
        {
            button3.ForeColor = Color.Green;
            checkBox1.Enabled = true;
            String name = "licschkre";
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            "2.xlsx" +
                            ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(constr);
            OleDbCommand oconn = new OleDbCommand("Select Adı, [Lizinq obyektin dəyəri], [Lizinqin məbləği], [Qalıq], [V#K#Qalıq], [% məbləği], [V#K#% məbləği], [Cərimə % məbləği], [Son əməl#tarixi], [Ö#G#], [Layihe], [Qrafikda olan məbləğ] From [" + name + "$] where Layihe like '%S-%' or Layihe like '%L-01/14-126/08%' or Layihe like '%01/13-106/08%' or Layihe like '%L-158%' or Layihe like '%L-079%' or Layihe like '%L-174%' or Layihe like '%01/08-105/08%' or Layihe like '%02/08-105/08%' or Layihe like '%L-198%' or Layihe like '%L-193%' or Layihe like '%L-161%' or Layihe like '%L-196%' or Layihe like '%L-114%' or Layihe like '%01/14-167/10%' or Layihe like '%L-02/14-167-10%' or Layihe like '%03/14-167/10%' or Layihe like '%04/16-167/10%' or Layihe like '%L-195/11%' or Layihe like '%L-01/16%'", con);
            con.Open();

            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
            DataTable data = new DataTable();
            data.Clear();
            sda.Fill(data);
            dataGridView1.DataSource = data;
            con.Close();

            label1.Text = dataGridView1.Rows.Count.ToString();

            checkBox1.Checked = false;
            checkBox1.Checked = true;
        }

        private static CultureInfo ci = new CultureInfo("AZ");

        private void txtaxtar_KeyDown(object sender, KeyEventArgs e)
        {
            if (radioLizinq.Checked == true)  //eger Lizinq uzre axtaririqsa ------------------------------------------------------------------
            {
                try
                {
                    if (e.KeyCode == Keys.Enter)
                    {
                        txtaxtar.Text = txtaxtar.Text.ToUpper(ci);

                        String name = "licschkre";
                        String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                        "2.xlsx" +
                                        ";Extended Properties='Excel 12.0 xml;HDR=YES;';";
                        
                        OleDbConnection con = new OleDbConnection(constr);
                        OleDbCommand oconn = new OleDbCommand("Select Adı, [Lizinq obyektin dəyəri], [Lizinqin məbləği], [Qalıq], [V#K#Qalıq], [% məbləği], [V#K#% məbləği], [Cərimə % məbləği], [Son əməl#tarixi], [Ö#G#], [Layihe], [Qrafikda olan məbləğ] From [" + name + "$] WHERE Adı Like '%" + txtaxtar.Text + "%' or Layihe Like '%" + txtaxtar.Text + "%'", con);
                        con.Open();

                        OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                        DataTable data = new DataTable();
                        data.Clear();
                        sda.Fill(data);
                        dataGridView1.DataSource = data;
                        con.Close();

                        /*BindingSource bs = new BindingSource();
                        bs.DataSource = dataGridView1.DataSource;
                        bs.Filter = "Adı like '%" + txtaxtar.Text + "%'";
                        dataGridView1.DataSource = bs;
                        con.Close();

                        dataGridView1.Columns.Remove("Reqion");
                        dataGridView1.Columns.Remove(" ");
                        dataGridView1.Columns.Remove(" 1");
                        dataGridView1.Columns.Remove("İnd#");
                        dataGridView1.Columns.Remove("K#S#");
                        dataGridView1.Columns.Remove("Lizinq hesabı");
                        dataGridView1.Columns.Remove("Faiz hesabı");
                        dataGridView1.Columns.Remove("Lizinq V#K#hesabı");
                        dataGridView1.Columns.Remove("Faiz");
                        dataGridView1.Columns.Remove("V#K#Faiz");
                        dataGridView1.Columns.Remove("Ehtiyat");
                        dataGridView1.Columns.Remove("V#K#Eht#");
                        dataGridView1.Columns.Remove("Bir ayın % məbləği");
                        dataGridView1.Columns.Remove("Avans");
                        dataGridView1.Columns.Remove("Sığorta");
                        dataGridView1.Columns.Remove("Servis");
                        dataGridView1.Columns.Remove("Dəbbə");
                        dataGridView1.Columns.Remove("Bağlanma");
                        dataGridView1.Columns.Remove("U#S#");
                        dataGridView1.Columns.Remove("Uzadılma tarixi");
                        dataGridView1.Columns.Remove("Verilmə tarixi");
                        dataGridView1.Columns.Remove("L#K#");
                        dataGridView1.Columns.Remove("L#M#");
                        dataGridView1.Columns.Remove("Lizinqin növü");
                        dataGridView1.Columns.Remove("Girovun növü");
                        dataGridView1.Columns.Remove("Girovun hesabı");
                        dataGridView1.Columns.Remove("Girovun məbləği");
                        dataGridView1.Columns.Remove("Y#Q#S#");
                        dataGridView1.Columns.Remove("Girovun yeni qiyməti");
                        dataGridView1.Columns.Remove("Y#Q#T#");
                        dataGridView1.Columns.Remove("Lizinq mənbələrin növü");
                        dataGridView1.Columns.Remove("Val#");
                        dataGridView1.Columns.Remove("K#T#");
                        dataGridView1.Columns.Remove("G#M#");
                        dataGridView1.Columns.Remove("Tranzit hesabı");
                        dataGridView1.Columns.Remove("V#K# tranzit hesabı");
                        dataGridView1.Columns.Remove("overdueday");
                        dataGridView1.Columns.Remove("Dəbbə %");
                        dataGridView1.Columns.Remove("Alqı-satqı müqaviləsinin nömrəsi");
                        dataGridView1.Columns.Remove("Satıcı");
                        dataGridView1.Columns.Remove("Təhvil-təslim aktı - tarix");
                        dataGridView1.Columns.Remove("Note");
                        dataGridView1.Columns.Remove("SIGORTA_UZRE_MESUL_SHEXS");
                        dataGridView1.Columns.Remove("SIGORTA_POLISININ_SERIYASI");
                        dataGridView1.Columns.Remove("SIGORTA_POLISININ_NOMRESI");
                        dataGridView1.Columns.Remove("SIGORTA_POLISININ_VER_TARIXI");
                        dataGridView1.Columns.Remove("SIGORTA_POLISININ_BIT_TARIXI");
                        dataGridView1.Columns.Remove("SIGORTALANAN_MEBLEG");
                        dataGridView1.Columns.Remove("SIGORTA_SHIRKET_HESAB_NOMRESI");
                        dataGridView1.Columns.Remove("TYPE OF SALE");
                        dataGridView1.Columns.Remove("CLIENT");
                        dataGridView1.Columns.Remove("STATUS");*/
                    }
                }
                catch { MessageBox.Show("Axtarış alınmadı"); };
            }


            //eger bank uzre axtaririqsa --------------------------------------------------------------------------------------
            if (radioBank.Checked == true)
            {
                try
                {
                    if (e.KeyCode == Keys.Enter)
                    {
                        txtaxtar.Text = txtaxtar.Text.ToUpper(ci);

                        String name = "Кредитный портфель";
                        String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                        "1.xlsx" +
                                        ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

                        OleDbConnection con = new OleDbConnection(constr);
                        OleDbCommand oconn = new OleDbCommand("Select [ID клиента], [Номер контракта], [Наименование клиента], [Сумма контракта в манатном эквиваленте], [Остаток основного долга на дату в манатном эквиваленте], [Остаток просрочки на дату в манатном эквиваленте], [Просроченные проценты в манатном эквиваленте], [Штрафные проценты в манатном эквиваленте], [Начисленные проценты в манатном эквиваленте], [Регион кредита], [Филиал], [Дата заключения контракта], [Дата окончания контракта], [Валюта контракта], [Срок просрочки процентов], [Последняя дата погашения процентов (или просроченных процентов)] From [" + name + "$] WHERE [ID клиента] Like '%" + txtaxtar.Text + "%' or [Наименование клиента] Like '%" + txtaxtar.Text + "%' or [Филиал] Like '%" + txtaxtar.Text + "%'", con);
                        con.Open();

                        OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                        DataTable data = new DataTable();
                        data.Clear();
                        sda.Fill(data);
                        dataGridView1.DataSource = data;
                        con.Close();

                        
                    }
                }
                catch { };
            }

            try { label1.Text = dataGridView1.Rows.Count.ToString(); } catch { }
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.ForeColor = Color.Green;
            button2.ForeColor = Color.Red;
            button3.ForeColor = Color.Red;
            button4.ForeColor = Color.Red;
            button5.ForeColor = Color.Red;
            button6.ForeColor = Color.Red;
            button7.ForeColor = Color.Red; 
            
            checkBox1.Enabled = true;
            String name = "licschkre";
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            "2.xlsx" +
                            ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(constr);
            OleDbCommand oconn = new OleDbCommand("Select Adı, [Lizinq obyektin dəyəri], [Lizinqin məbləği], [Qalıq], [V#K#Qalıq], [% məbləği], [V#K#% məbləği], [Cərimə % məbləği], [Son əməl#tarixi], [Ö#G#], [Layihe], [Qrafikda olan məbləğ] From [" + name + "$] where Layihe like '%L%'", con);
            con.Open();

            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
            DataTable data = new DataTable();
            data.Clear();
            sda.Fill(data);
            dataGridView1.DataSource = data;
            con.Close();

            label1.Text = dataGridView1.Rows.Count.ToString();
            checkBox1.Checked = false;
            checkBox1.Checked = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            button1.ForeColor = Color.Red;
            button2.ForeColor = Color.Green;
            button3.ForeColor = Color.Red;
            button4.ForeColor = Color.Red;
            button5.ForeColor = Color.Red;
            button6.ForeColor = Color.Red;
            button7.ForeColor = Color.Red;


            checkBox1.Enabled = true;
            String name = "licschkre";
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            "2.xlsx" +
                            ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(constr);
            OleDbCommand oconn = new OleDbCommand("Select Adı, [Lizinq obyektin dəyəri], [Lizinqin məbləği], [Qalıq], [V#K#Qalıq], [% məbləği], [V#K#% məbləği], [Cərimə % məbləği], [Son əməl#tarixi], [Ö#G#], [Layihe], [Qrafikda olan məbləğ] From [" + name + "$] where Layihe like '%S%'", con);
            con.Open();

            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
            DataTable data = new DataTable();
            data.Clear();
            sda.Fill(data);
            dataGridView1.DataSource = data;
            con.Close();

            label1.Text = dataGridView1.Rows.Count.ToString();

            checkBox1.Checked = false;
            checkBox1.Checked = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            button1.ForeColor = Color.Red;
            button2.ForeColor = Color.Red;
            button3.ForeColor = Color.Green;
            button4.ForeColor = Color.Red;
            button5.ForeColor = Color.Red;
            button6.ForeColor = Color.Red;
            button7.ForeColor = Color.Red; 
            
            checkBox1.Enabled = true;
            String name = "licschkre";
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            "2.xlsx" +
                            ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(constr);
            // [V.K.% məbləği], [K.p.b.tarixi], [Son əməl.tarixi], [Ö.G.], [Layihe], [Qrafikda olan məbləğ] 
            OleDbCommand oconn = new OleDbCommand("Select Adı, [Lizinq obyektin dəyəri], [Lizinqin məbləği], [Qalıq], [V#K#Qalıq], [% məbləği], [V#K#% məbləği], [Cərimə % məbləği], [Son əməl#tarixi], [Ö#G#], [Layihe], [Qrafikda olan məbləğ] From [" + name + "$] where Layihe like '%S-%' or Layihe like '%L-01/14-126/08%' or Layihe like '%01/13-106/08%' or Layihe like '%L-158%' or Layihe like '%L-079%' or Layihe like '%L-174%' or Layihe like '%01/08-105/08%' or Layihe like '%02/08-105/08%' or Layihe like '%L-198%' or Layihe like '%L-193%' or Layihe like '%L-161%' or Layihe like '%L-196%' or Layihe like '%L-114%' or Layihe like '%01/14-167/10%' or Layihe like '%L-02/14-167-10%' or Layihe like '%03/14-167/10%' or Layihe like '%04/16-167/10%' or Layihe like '%L-195/11%' or Layihe like '%L-01/16%'", con);
            con.Open();

            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
            DataTable data = new DataTable();
            data.Clear();
            sda.Fill(data);
            dataGridView1.DataSource = data;
            con.Close();

            label1.Text = dataGridView1.Rows.Count.ToString();

            checkBox1.Checked = false;
            checkBox1.Checked = true;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            button1.ForeColor = Color.Red;
            button2.ForeColor = Color.Red;
            button3.ForeColor = Color.Red;
            button4.ForeColor = Color.Green;
            button5.ForeColor = Color.Red;
            button6.ForeColor = Color.Red;
            button7.ForeColor = Color.Red;

            checkBox1.Enabled = true; 
            String name = "licschkre";
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            "2.xlsx" +
                            ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(constr);
            OleDbCommand oconn = new OleDbCommand("Select Adı, [Lizinq obyektin dəyəri], [Lizinqin məbləği], [Qalıq], [V#K#Qalıq], [% məbləği], [V#K#% məbləği], [Cərimə % məbləği], [Son əməl#tarixi], [Ö#G#], [Layihe], [Qrafikda olan məbləğ] From [" + name + "$]", con);
            con.Open();

            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
            DataTable data = new DataTable();
            data.Clear();
            sda.Fill(data);
            dataGridView1.DataSource = data;
            con.Close();

            label1.Text = dataGridView1.Rows.Count.ToString();

            checkBox1.Checked = false;
            checkBox1.Checked = true;
        }

        private void printToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int s, k, a=0;
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;

            try { File.Copy("Portfel.xlsm", "C:\\Users\\" + Environment.UserName + "\\Desktop\\PORTFEL.xlsm", true); }
            catch { MessageBox.Show("'Portfel.xlsm' tapılmadı."); }

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\PORTFEL.xlsm"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];

            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

            oSheet.Cells.EntireColumn.NumberFormat = "@";

            if (checkBox1.Checked == false)
            {
                for (s = 0; s < dataGridView1.Columns.Count; s++)
                {
                    oSheet.Cells[1, s + 2] = dataGridView1.Columns[s].HeaderText;

                    for (k = 0; k < dataGridView1.Rows.Count; k++)
                    {
                        oSheet.Cells[k + 2, s + 2] = dataGridView1.Rows[k].Cells[s].Value;
                        //oSheet.Cells[k + 1, s + 1].Borders.LineStyle = Excel.Constants.xlSolid;

                        if (k != dataGridView1.Rows.Count - 1) oSheet.Cells[k + 2, 1] = k + 1;
                    }
                    
                    oSheet.Columns.AutoFit();
                    oSheet.Rows.AutoFit();
                }
            }

            if (checkBox1.Checked == true )
            {
                for (s = 0; s < dataGridView1.Columns.Count; s++)
                {
                    if (s == 0) a = 5;
                    if (s == 1) a = 6;
                    if (s == 2) a = 6;
                    if (s == 3) a = 7;
                    if (s == 4) a = 8;
                    if (s == 5) a = 9;
                    if (s == 6) a = 10;
                    if (s == 7) a = 11;
                    if (s == 8) a = 12;
                    if (s == 9) a = 2;
                    if (s == 10) a = 4; 
                    if (s == 11) a = 13;

                    oSheet.Cells[1, a] = dataGridView1.Columns[s].HeaderText;
                    oSheet.Cells[1, 1] = "Nö";

                    for (k = 0; k < dataGridView1.Rows.Count; k++)
                    {
                        oSheet.Cells[k + 2, a] = dataGridView1.Rows[k].Cells[s].Value.ToString();
                        //oSheet.Cells[k + 1, s + 1].Borders.LineStyle = Excel.Constants.xlSolid;
                        
                        if (k != dataGridView1.Rows.Count - 1) oSheet.Cells[k + 2, 1] = k + 1;
                    }
                    oSheet.Columns.AutoFit();
                    oSheet.Rows.AutoFit();
                }

            }

            oWB.Save();

            //  oSheet.PrintOut();
            //  oWB.Close(SaveChanges: false);
            //  oXL.Application.Quit();

        }

        private void çıxışToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult result2 = MessageBox.Show("Davam etmək istəyirsiniz?", "?", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (result2 == DialogResult.No) { return; }

            base.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            button1.ForeColor = Color.Red;
            button2.ForeColor = Color.Red;
            button3.ForeColor = Color.Red;
            button4.ForeColor = Color.Red;
            button5.ForeColor = Color.Green;
            button6.ForeColor = Color.Red;
            button7.ForeColor = Color.Red; 
            
            checkBox1.Checked = false;
            checkBox1.Enabled = false;
            String name = "Кредитный портфель";
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            "1.xlsx" +
                            ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(constr);
            OleDbCommand oconn = new OleDbCommand("Select [ID клиента], [Наименование клиента], [Дата заключения контракта], [Дата окончания контракта], [Сумма контракта в манатном эквиваленте], [Остаток основного долга на дату в манатном эквиваленте], [Остаток просрочки на дату в манатном эквиваленте], [Начисленные проценты в манатном эквиваленте], [Просроченные проценты в манатном эквиваленте], [Штрафные проценты в манатном эквиваленте], [Срок просрочки процентов], [Последняя дата погашения процентов (или просроченных процентов)] From [" + name + "$]", con);
            con.Open();

            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
            DataTable data = new DataTable();
            data.Clear();
            sda.Fill(data);
            dataGridView1.DataSource = data;
            con.Close();

            label1.Text = dataGridView1.Rows.Count.ToString();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            button1.ForeColor = Color.Red;
            button2.ForeColor = Color.Red;
            button3.ForeColor = Color.Red;
            button4.ForeColor = Color.Red;
            button5.ForeColor = Color.Red;
            button6.ForeColor = Color.Green;
            button7.ForeColor = Color.Red;

            checkBox1.Checked = false;
            checkBox1.Enabled = false;
            String name = "Кредитный портфель";
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            "1.xlsx" +
                            ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(constr);
            OleDbCommand oconn = new OleDbCommand("Select [ID клиента], [Наименование клиента], [Дата заключения контракта], [Дата окончания контракта], [Сумма контракта в манатном эквиваленте], [Остаток основного долга на дату в манатном эквиваленте], [Остаток просрочки на дату в манатном эквиваленте], [Начисленные проценты в манатном эквиваленте], [Просроченные проценты в манатном эквиваленте], [Штрафные проценты в манатном эквиваленте], [Срок просрочки процентов], [Последняя дата погашения процентов (или просроченных процентов)] From [" + name + "$] WHERE [Способ выдачи кредита] like '%LEASING%'", con);
            con.Open();
            
            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
            DataTable data = new DataTable();
            data.Clear();
            sda.Fill(data);
            dataGridView1.DataSource = data;

            /*
            BindingSource bs = new BindingSource();
            bs.DataSource = dataGridView1.DataSource;
            bs.Filter = "[Способ выдачи кредита] like '%LEASING%'";
            dataGridView1.DataSource = bs;
            con.Close();

            for (i = 0; i < dataGridView1.Columns.Count; i++)
            {
                dataGridView1.Columns[i].Visible = false;
            }

            for (i = 0; i < dataGridView1.Columns.Count; i++)
            {
                if (dataGridView1.Columns[i].HeaderText == "ID клиента") dataGridView1.Columns[i].Visible = true;
                if (dataGridView1.Columns[i].HeaderText == "Наименование клиента") dataGridView1.Columns[i].Visible = true;
                if (dataGridView1.Columns[i].HeaderText == "Дата заключения контракта") dataGridView1.Columns[i].Visible = true;
                if (dataGridView1.Columns[i].HeaderText == "Дата окончания контракта") dataGridView1.Columns[i].Visible = true;
                if (dataGridView1.Columns[i].HeaderText == "Сумма контракта в манатном эквиваленте") dataGridView1.Columns[i].Visible = true;
                if (dataGridView1.Columns[i].HeaderText == "Остаток основного долга на дату в манатном эквиваленте") dataGridView1.Columns[i].Visible = true;
                if (dataGridView1.Columns[i].HeaderText == "Остаток просрочки на дату в манатном эквиваленте") dataGridView1.Columns[i].Visible = true;
                if (dataGridView1.Columns[i].HeaderText == "Начисленные проценты в манатном эквиваленте") dataGridView1.Columns[i].Visible = true;
                if (dataGridView1.Columns[i].HeaderText == "Просроченные проценты в манатном эквиваленте") dataGridView1.Columns[i].Visible = true;
                if (dataGridView1.Columns[i].HeaderText == "Штрафные проценты в манатном эквиваленте") dataGridView1.Columns[i].Visible = true;
                if (dataGridView1.Columns[i].HeaderText == "Срок просрочки процентов") dataGridView1.Columns[i].Visible = true;
                if (dataGridView1.Columns[i].HeaderText == "Последняя дата погашения процентов (или просроченных процентов)") dataGridView1.Columns[i].Visible = true;
            }

            try { dataGridView1.Columns.Remove("Проценты на внебалансе в манатном эквиваленте"); }
            catch { MessageBox.Show("Проценты на внебалансе в манатном эквиваленте tapılmadı"); };
            */
            label1.Text = dataGridView1.Rows.Count.ToString();
        }

        private void radioLizinq_CheckedChanged(object sender, EventArgs e)
        {
            checkBox1.Enabled = true;
        }

        private void radioBank_CheckedChanged(object sender, EventArgs e)
        {
            checkBox1.Checked = false;
            checkBox1.Enabled = false;
        }

        private void əlaqəToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Elaqe elaqe = new Elaqe();
            elaqe.ShowDialog();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            try { this.dataGridView1.Sort(this.dataGridView1.Columns["Layihe"], ListSortDirection.Descending); }
            catch { };
        }

        private void button7_Click_2(object sender, EventArgs e)
        {
            button1.ForeColor = Color.Red;
            button2.ForeColor = Color.Red;
            button3.ForeColor = Color.Red;
            button4.ForeColor = Color.Red;
            button5.ForeColor = Color.Red;
            button6.ForeColor = Color.Red;
            button7.ForeColor = Color.Green;

            checkBox1.Enabled = true;
            String name = "licschkre";
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            "2.xlsx" +
                            ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(constr);
            // [V.K.% məbləği], [K.p.b.tarixi], [Son əməl.tarixi], [Ö.G.], [Layihe], [Qrafikda olan məbləğ] 
            OleDbCommand oconn = new OleDbCommand("Select Adı, [Lizinq obyektin dəyəri], [Lizinqin məbləği], [Qalıq], [V#K#Qalıq], [% məbləği], [V#K#% məbləği], [Cərimə % məbləği], [Son əməl#tarixi], [Ö#G#], [Layihe], [Qrafikda olan məbləğ] From [" + name + "$] where Layihe like '%S-%' and [V#K#Qalıq] + [V#K#% məbləği] > 0", con);
            con.Open();

            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
            DataTable data = new DataTable();
            data.Clear();
            sda.Fill(data);
            dataGridView1.DataSource = data;
            con.Close();

            label1.Text = dataGridView1.Rows.Count.ToString();

            checkBox1.Checked = false;
            checkBox1.Checked = true;
            
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {

        }

        private void ödənişQəbziToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Qəbz qebz = new Qəbz();
            qebz.Show();
        }

        private void haqqındaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Help help = new Help();
            help.ShowDialog();
        }

    }
}
