using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Globalization;
using System.IO;
using System.Net;
using System.Web;
using System.Data.OleDb;
using Nsoft;

namespace Lizinq_Muqavile
{
    public partial class Tehvil_Teslim : Form
    {
        public Tehvil_Teslim()
        {
            InitializeComponent();
        }
        public void WordDoc()
        {
            try { File.Copy("Tehvil-Teslim.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Tehvil-Teslim - " + dttarix.Text + ".doc", true); }
            catch { MessageBox.Show("'X:\\AGLizinq\\Tehvil-Teslim.doc' tapılmadı."); }

            object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\Tehvil-Teslim - " + dttarix.Text + ".doc";

            string k = "", TVSobe = "", TASobe = "", TVqurum = "", TAqurum = "";
            if (rbesli.Checked == true) k = rbesli.Text;
            if (rbsureti.Checked == true) k = rbsureti.Text;


            Word.Application word = new Word.Application();
            Word.Document doc = null;
            object missing = System.Type.Missing;
            object readOnly = false;
            object isVisible = false;
            word.Visible = true;

            doc = word.Documents.Open(ref FileName);
            doc.Activate();

            DateTime dt = dttarix.Value.Date;
            string tarix = MyChange.TarixSozle(dt);

            MyChange.FindAndReplace(word, "000", dt.Day + " " + tarix + " " + dt.Year + "-ci il");
            if (cbsobeTehvilVeren.Text.Substring(1, 8) == "AGLizinq" && rbesli.Checked == true) MyChange.FindAndReplace(word, "111", "Mən, “AGLizinq” QSC-nin " + cbvezifeTehvilVeren.Text + "i " + cbTehvilVerenMutexessis.Text + ", " + cbsobeTehvilAlan.Text + "-nin " + cbsobeTehvilAlan2.Text + "nə,  aşağıda göstərilən lizinq layihəsinin əslini təhvil verdim: ");
            if (cbsobeTehvilVeren.Text.Substring(1, 8) == "AGLizinq" && rbsureti.Checked == true) MyChange.FindAndReplace(word, "111", "Mən, “AGLizinq” QSC-nin " + cbvezifeTehvilVeren.Text + "i " + cbTehvilVerenMutexessis.Text + ", " + cbsobeTehvilAlan.Text + "-nin " + cbsobeTehvilAlan2.Text + "nə,  aşağıda göstərilən lizinq layihəsinin surətini təhvil verdim: ");
            if (cbsobeTehvilVeren.Text.Substring(1, 8) != "AGLizinq" && rbesli.Checked == true) MyChange.FindAndReplace(word, "111", "Mən, " + cbsobeTehvilVeren.Text + "-nin " + cbsobeTehvilVeren2.Text + "nin " + cbvezifeTehvilVeren.Text + "i " + cbTehvilVerenMutexessis.Text + ", " + cbsobeTehvilAlan.Text + "-nə,  aşağıda göstərilən lizinq layihəsinin əslini təhvil verdim: ");
            if (cbsobeTehvilVeren.Text.Substring(1, 8) != "AGLizinq" && rbsureti.Checked == true) MyChange.FindAndReplace(word, "111", "Mən, " + cbsobeTehvilVeren.Text + "-nin " + cbsobeTehvilVeren2.Text + "nin " + cbvezifeTehvilVeren.Text + "i " + cbTehvilVerenMutexessis.Text + ", " + cbsobeTehvilAlan.Text + "-nə,  aşağıda göstərilən lizinq layihəsinin surətini təhvil verdim: ");

            if (t1.Text != "" && t2.Text != "") MyChange.FindAndReplace(word, "222", t2.Text + " " + t1.Text + " layihəsi."); else MyChange.FindAndReplace(word, "222", "");
            if (t3.Text != "" && t4.Text != "") MyChange.FindAndReplace(word, "333", t4.Text + " " + t3.Text + " layihəsi."); else MyChange.FindAndReplace(word, "333", "");
            if (t5.Text != "" && t6.Text != "") MyChange.FindAndReplace(word, "444", t6.Text + " " + t5.Text + " layihəsi."); else MyChange.FindAndReplace(word, "444", "");

            try
            {
                if (cbsobeTehvilVeren.Text.Substring(1, 6) == "AGBank") { TVSobe = "“AGBank” ASC"; MyChange.FindAndReplace(word, "7777", "AGBank ASC"); MyChange.FindAndReplace(word, "8888", "AGLizinq QSC"); }
                else { TVSobe = cbsobeTehvilVeren.Text; MyChange.FindAndReplace(word, "7777", "AGLizinq QSC"); MyChange.FindAndReplace(word, "8888", "AGBank ASC"); }
            }
            catch { }

            try
            {
                TVqurum = cbsobeTehvilVeren2.Text;
                if (cbsobeTehvilVeren2.Text.Substring(0, 1) == "M") { TVqurum = "Hüquq"; }
                else if (cbsobeTehvilVeren2.Text.Substring(0, 1) == " ") { TVqurum = "Lizinq"; }
                else if (cbsobeTehvilVeren2.Text.Substring(0, 1) == "P") { TVqurum = "Problemi"; }
            }
            catch { }

            try
            {
                if (cbsobeTehvilAlan.Text.Substring(1, 6) == "AGBank") { TASobe = "“AGBank” ASC"; }
                else { TASobe = cbsobeTehvilAlan.Text; }
            }
            catch { }

            try
            {
                TAqurum = cbsobeTehvilAlan2.Text;
                if (cbsobeTehvilAlan2.Text.Substring(0, 1) == " ") { TAqurum = "Lizinq"; }
                else if (cbsobeTehvilAlan2.Text.Substring(0, 1) == "M") { TAqurum = "Hüquq"; }
                else if (cbsobeTehvilAlan2.Text.Substring(0, 1) == "P") { TAqurum = "Problemli"; }
            }
            catch { }

            try { MyChange.FindAndReplace(word, "888", cbTehvilVerenMutexessis.Text); }
            catch { }
            MyChange.FindAndReplace(word, "999", cbTehvilAlanMutexessis.Text);

            if (t1.Text != "")
            {
                try
                {
                    MyData.insertCommand("baza.accdb", "insert into TehvilTeslim (a1, a2, a3, a4, a5, a6) Values ('" + t2.Text + "', '" + t1.Text + "', '" + dt.Day + " " + tarix + " " + dt.Year + "', '" + TVSobe + "/" + TVqurum + "/" + cbTehvilVerenMutexessis.Text + "','" + TASobe + "/" + TAqurum + "/" + cbTehvilAlanMutexessis.Text + "', '" + txtqeydler.Text + "')");
                }
                catch { MessageBox.Show("Tehvil Teslim bazaya qeyd olunarkən sehv..."); }

                try
                {
                    MyData.insertCommand("baza.accdb", "insert into Emeliyyatlar (a1, a2, a3) Values ('" + dt.Date + "', 'TƏHVİL TƏSLİM - " + t1.Text + " (" + t2.Text + ") - " + cbTehvilAlanMutexessis.Text + " (" + cbsobeTehvilAlan.Text + ") " + k + "','" + Environment.MachineName + "')");
                }
                catch { MessageBox.Show("Tehvil Teslim Emeliyyatlara qeyd olunarkən sehv..."); }

                try
                {
                    MyData.updateCommand("baza.accdb", "UPDATE PortfelStatus SET "
                                                                                         + "Adı ='" + t1.Text + "',"
                                                                                         + "Layihe ='" + t2.Text + "',"
                                                                                         + "Status ='" + cbsobeTehvilAlan2.Text.Substring(0, 1).ToString() + "'"
                                                                                         + " WHERE Layihe Like '%" + t2.Text + "%'");
                }
                catch { }

            }

            if (t3.Text != "")
            {
                try
                {
                    MyData.insertCommand("baza.accdb", "insert into TehvilTeslim (a1, a2, a3, a4, a5, a6) Values ('" + t4.Text + "', '" + t3.Text + "', '" + dttarix.Text.Substring(0, 2) + " " + tarix + " " + dttarix.Text.Substring(dttarix.Text.Length - 4, 4) + "', '" + TVSobe + "/" + TVqurum + "/" + cbTehvilVerenMutexessis.Text + "','" + TASobe + "/" + TAqurum + "/" + cbTehvilAlanMutexessis.Text + "', '" + txtqeydler.Text + "')");
                }
                catch { MessageBox.Show("Tehvil Teslim bazaya qeyd olunarkən sehv..."); }

                try
                {
                    MyData.insertCommand("baza.accdb", "insert into Emeliyyatlar (a1, a2, a3) Values ('" + dt.Date + "', 'TƏHVİL TƏSLİM - " + t3.Text + " (" + t4.Text + ") - " + cbTehvilAlanMutexessis.Text + " (" + cbsobeTehvilAlan.Text + ") " + k + "','" + Environment.MachineName + "')");
                }
                catch { MessageBox.Show("Tehvil Teslim Emeliyyatlara qeyd olunarkən sehv..."); }

                try
                {
                    MyData.updateCommand("baza.accdb", "UPDATE PortfelStatus SET "
                                                                                            + "Adı ='" + t3.Text + "',"
                                                                                         + "Layihe ='" + t4.Text + "',"
                                                                                         + "Status ='" + cbsobeTehvilAlan2.Text.Substring(0, 1).ToString() + "'"
                                                                                         + " WHERE Layihe Like '%" + t4.Text + "%'");
                }
                catch { }

            }

            if (t5.Text != "")
            {
                try
                {
                    MyData.insertCommand("baza.accdb", "insert into TehvilTeslim (a1, a2, a3, a4, a5, a6) Values ('" + t6.Text + "', '" + t5.Text + "', '" + dttarix.Text.Substring(0, 2) + " " + tarix + " " + dttarix.Text.Substring(dttarix.Text.Length - 4, 4) + "', '" + TVSobe + "/" + TVqurum + "/" + cbTehvilVerenMutexessis.Text + "','" + TASobe + "/" + TAqurum + "/" + cbTehvilAlanMutexessis.Text + "', '" + txtqeydler.Text + "')");
                }
                catch { MessageBox.Show("Tehvil Teslim bazaya qeyd olunarkən sehv..."); }

                try
                {
                    MyData.insertCommand("baza.accdb", "insert into Emeliyyatlar (a1, a2, a3) Values ('" + dt.Date + "', 'TƏHVİL TƏSLİM - " + t5.Text + " (" + t6.Text + ") - " + cbTehvilAlanMutexessis.Text + " (" + cbsobeTehvilAlan.Text + ") " + k + "','" + Environment.MachineName + "')");
                }
                catch { MessageBox.Show("Tehvil Teslim Emeliyyatlara qeyd olunarkən sehv..."); }

                try
                {
                    MyData.updateCommand("baza.accdb", "UPDATE PortfelStatus SET "
                                                                                         + "Adı ='" + t5.Text + "',"
                                                                                         + "Layihe ='" + t6.Text + "',"
                                                                                         + "Status ='" + cbsobeTehvilAlan2.Text.Substring(0, 1).ToString() + "'"
                                                                                         + " WHERE Layihe Like '%" + t6.Text + "%'");
                }
                catch { }

            }
        }

        private static CultureInfo ci = new CultureInfo("AZ");

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult result2 = MessageBox.Show("Davam etmək istəyirsiniz?", "?", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (result2 == DialogResult.No) { return; }

            WordDoc();
        }

        private void cbtehvilveren_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbTehvilVerenMutexessis.Text == "Əsədov İbrahim Akif oğlu") cbvezifeTehvilVeren.Text = "Filial müdiri";
            else if (cbTehvilVerenMutexessis.Text == "Musayev Rəşad İslam oğlu") cbvezifeTehvilVeren.Text = "Baş menecer";
            else if (cbTehvilVerenMutexessis.Text == "Heydərov Namik Abisalam oğlu") cbvezifeTehvilVeren.Text = "Baş mütəxəssis";
            else if (cbTehvilVerenMutexessis.Text == "Şirinov Rasim Rafiqoviç") cbvezifeTehvilVeren.Text = "Sürücü";
            else if (cbTehvilVerenMutexessis.Text == "Elqar Ələkbərov") cbvezifeTehvilVeren.Text = "Baş hüquqşünas";

        }

        private void cbsobe_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbsobeTehvilAlan.Text.Substring(1, 8) == "AGLizinq") cbsobeTehvilAlan2.Text = "";
            if (cbsobeTehvilAlan2.Text == "Problemli Aktivlərin İdarə Edilməsi Departamenti") cbTehvilAlanMutexessis.Text = "Məcidov Asim";
            else if (cbsobeTehvilAlan2.Text == "Məhkəmə və İcra işləri Departamenti") cbTehvilAlanMutexessis.Text = "Məmmədova Dilarə";
        }

        private void t1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                t2.Text = "";

                try
                {
                    string commandText = "SELECT * FROM etibarnameneqliyyat WHERE 1=1";
                    try
                    {
                        commandText += " and c3 like '%" + t1.Text + "%'";
                    }
                    catch { };

                    MyData.selectCommand("baza.accdb", commandText);
                    MyData.dtmain = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmain);

                    try
                    {
                        t2.Text = MyData.dtmain.Rows[0]["c4"].ToString();
                        t1.Text = MyData.dtmain.Rows[0]["c3"].ToString();
                    }
                    catch { }


                    if (t2.Text == "")
                    {
                        String name = "licschkre";
                        String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                        "2.xlsx" +
                                        ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

                        OleDbConnection con = new OleDbConnection(constr);
                        OleDbCommand oconn = new OleDbCommand("Select Layihe, Adı From [" + name + "$] where Adı Like '%" + t1.Text + "%'", con);
                        con.Open();
                        OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                        DataTable data = new DataTable();
                        data.Clear();
                        sda.Fill(data);
                        con.Close();

                        try
                        {
                            t2.Text = data.Rows[0]["Layihe"].ToString();
                            t1.Text = data.Rows[0]["Adı"].ToString();
                            //t1.Text = t1.Text.Substring(0, 1).ToUpper(ci) + t1.Text.Substring(1, t1.Text.Length - 1).ToLower(ci);
                        }
                        catch { }

                    }

                    if (t2.Text == "")
                    {
                        String name2 = "Кредитный портфель";
                        String constr2 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                        "1.xlsx" +
                                        ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

                        OleDbConnection con2 = new OleDbConnection(constr2);
                        OleDbCommand oconn2 = new OleDbCommand("Select [Наименование клиента], [Номер контракта] From [" + name2 + "$] WHERE [Наименование клиента] Like '%" + t1.Text + "%'", con2);
                        con2.Open();
                        OleDbDataAdapter sda2 = new OleDbDataAdapter(oconn2);
                        DataTable data2 = new DataTable();
                        data2.Clear();
                        sda2.Fill(data2);
                        con2.Close();

                        try
                        {
                            t2.Text = data2.Rows[0]["Номер контракта"].ToString();
                            t1.Text = data2.Rows[0]["Наименование клиента"].ToString();
                        }
                        catch { }
                    }
                }
                catch { }

                try
                {
                    if (t1.Text.Substring(0, 1) == "'") t1.Text = t1.Text.Substring(1, t1.Text.Length - 1);
                    if (t2.Text.Substring(0, 1) == "'") t2.Text = t2.Text.Substring(1, t2.Text.Length - 1);
                }
                catch { }

            }
        }

        private void t2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                t1.Text = "";

                try
                {
                    MyData.selectCommand("baza.accdb", "SELECT * FROM etibarnameneqliyyat WHERE c4 like '%" + t2.Text + "%'");
                }
                catch { }

                MyData.dtmain = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmain);

                try
                {
                    t2.Text = MyData.dtmain.Rows[0]["c4"].ToString();
                    t1.Text = MyData.dtmain.Rows[0]["c3"].ToString();
                }
                catch { }

                if (t1.Text == "")
                {
                    String name = "licschkre";
                    String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    "2.xlsx" +
                                    ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

                    OleDbConnection con = new OleDbConnection(constr);
                    OleDbCommand oconn = new OleDbCommand("Select Layihe, Adı From [" + name + "$] where Layihe Like '%" + t2.Text + "%'", con);
                    con.Open();
                    OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                    DataTable data = new DataTable();
                    data.Clear();
                    sda.Fill(data);
                    con.Close();

                    try
                    {
                        t2.Text = data.Rows[0]["Layihe"].ToString();
                        t1.Text = data.Rows[0]["Adı"].ToString();
                        //t1.Text = t1.Text.Substring(0, 1).ToUpper(ci) + t1.Text.Substring(1, t1.Text.Length - 1).ToLower(ci);
                    }
                    catch { }
                }

                if (t1.Text == "")
                {
                    String name2 = "Кредитный портфель";
                    String constr2 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    "1.xlsx" +
                                    ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

                    OleDbConnection con2 = new OleDbConnection(constr2);
                    OleDbCommand oconn2 = new OleDbCommand("Select [Наименование клиента], [Номер контракта] From [" + name2 + "$] WHERE [Номер контракта] Like '%" + t2.Text + "%'", con2);
                    con2.Open();
                    OleDbDataAdapter sda2 = new OleDbDataAdapter(oconn2);
                    DataTable data2 = new DataTable();
                    data2.Clear();
                    sda2.Fill(data2);
                    con2.Close();

                    try
                    {
                        t2.Text = data2.Rows[0]["Номер контракта"].ToString();
                        t1.Text = data2.Rows[0]["Наименование клиента"].ToString();
                    }
                    catch { }
                }

                try
                {
                    if (t1.Text.Substring(0, 1) == "'") t1.Text = t1.Text.Substring(1, t1.Text.Length - 1);
                    if (t2.Text.Substring(0, 1) == "'") t2.Text = t2.Text.Substring(1, t2.Text.Length - 1);
                }
                catch { }
            }
        }
        private void btMMX2_Click(object sender, EventArgs e)
        {
            TehvilTeslim tehvilteslim = new TehvilTeslim();
            tehvilteslim.Show();
        }

        private void cbsobeTehvilVeren_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbsobeTehvilVeren.Text.Substring(1, 8) == "AGLizinq") cbsobeTehvilVeren2.Text = "";
        }

        private void t3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                t4.Text = "";

                string commandText = "SELECT * FROM etibarnameneqliyyat WHERE c3 like " + "'%" + t3.Text + "%'";

                MyData.selectCommand("baza.accdb", commandText);
                MyData.dtmain = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmain);

                try
                {
                    t4.Text = MyData.dtmain.Rows[0]["c4"].ToString();
                    t3.Text = MyData.dtmain.Rows[0]["c3"].ToString();
                }
                catch { }


                if (t4.Text == "")
                {
                    String name = "licschkre";
                    String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    "2.xlsx" +
                                    ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

                    OleDbConnection con = new OleDbConnection(constr);
                    OleDbCommand oconn = new OleDbCommand("Select Layihe, Adı From [" + name + "$] where Adı Like '%" + t3.Text + "%'", con);
                    con.Open();
                    OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                    DataTable data = new DataTable();
                    data.Clear();
                    sda.Fill(data);
                    con.Close();

                    try
                    {
                        t4.Text = data.Rows[0]["Layihe"].ToString();
                        t3.Text = data.Rows[0]["Adı"].ToString();
                        //t3.Text = t3.Text.Substring(0, 1).ToUpper(ci) + t3.Text.Substring(1, t3.Text.Length - 1).ToLower(ci);
                    }
                    catch { }

                }

                if (t4.Text == "")
                {
                    String name2 = "Кредитный портфель";
                    String constr2 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    "1.xlsx" +
                                    ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

                    OleDbConnection con2 = new OleDbConnection(constr2);
                    OleDbCommand oconn2 = new OleDbCommand("Select [Наименование клиента], [Номер контракта] From [" + name2 + "$] WHERE [Наименование клиента] Like '%" + t3.Text + "%'", con2);
                    con2.Open();
                    OleDbDataAdapter sda2 = new OleDbDataAdapter(oconn2);
                    DataTable data2 = new DataTable();
                    data2.Clear();
                    sda2.Fill(data2);
                    con2.Close();

                    try
                    {
                        t4.Text = data2.Rows[0]["Номер контракта"].ToString();
                        t3.Text = data2.Rows[0]["Наименование клиента"].ToString();
                    }
                    catch { }

                    try
                    {
                        if (t3.Text.Substring(0, 1) == "'") t3.Text = t3.Text.Substring(1, t3.Text.Length - 1);
                        if (t4.Text.Substring(0, 1) == "'") t4.Text = t4.Text.Substring(1, t4.Text.Length - 1);
                    }
                    catch { }
                }
            }
        }
        private void t5_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                t6.Text = "";

                string commandText = "SELECT * FROM etibarnameneqliyyat WHERE c3 like '%" + t5.Text + "%'";

                MyData.selectCommand("baza.accdb", commandText);
                MyData.dtmain = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmain);

                try
                {
                    t6.Text = MyData.dtmain.Rows[0]["c4"].ToString();
                    t5.Text = MyData.dtmain.Rows[0]["c3"].ToString();
                }
                catch { }


                if (t6.Text == "")
                {
                    String name = "licschkre";
                    String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    "2.xlsx" +
                                    ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

                    OleDbConnection con = new OleDbConnection(constr);
                    OleDbCommand oconn = new OleDbCommand("Select Layihe, Adı From [" + name + "$] where Adı Like '%" + t5.Text + "%'", con);
                    con.Open();
                    OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                    DataTable data = new DataTable();
                    data.Clear();
                    sda.Fill(data);
                    con.Close();

                    try
                    {
                        t6.Text = data.Rows[0]["Layihe"].ToString();
                        t5.Text = data.Rows[0]["Adı"].ToString();
                        //t5.Text = t5.Text.Substring(0, 1).ToUpper(ci) + t5.Text.Substring(1, t5.Text.Length - 1).ToLower(ci);
                    }
                    catch { }
                }
                if (t6.Text == "")
                {
                    String name2 = "Кредитный портфель";
                    String constr2 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    "1.xlsx" +
                                    ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

                    OleDbConnection con2 = new OleDbConnection(constr2);
                    OleDbCommand oconn2 = new OleDbCommand("Select [Наименование клиента], [Номер контракта] From [" + name2 + "$] WHERE [Наименование клиента] Like '%" + t5.Text + "%'", con2);
                    con2.Open();
                    OleDbDataAdapter sda2 = new OleDbDataAdapter(oconn2);
                    DataTable data2 = new DataTable();
                    data2.Clear();
                    sda2.Fill(data2);
                    con2.Close();

                    try
                    {
                        t6.Text = data2.Rows[0]["Номер контракта"].ToString();
                        t5.Text = data2.Rows[0]["Наименование клиента"].ToString();
                    }
                    catch { }

                    try
                    {
                        if (t5.Text.Substring(0, 1) == "'") t5.Text = t5.Text.Substring(1, t5.Text.Length - 1);
                        if (t6.Text.Substring(0, 1) == "'") t6.Text = t6.Text.Substring(1, t6.Text.Length - 1);
                    }
                    catch { }
                }
            }
        }
        private void t6_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                t5.Text = "";
                MyData.selectCommand("baza.accdb", "SELECT * FROM etibarnameneqliyyat WHERE c4 like " + "'%" + t6.Text + "%'");

                MyData.dtmain = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmain);

                try
                {
                    t6.Text = MyData.dtmain.Rows[0]["c4"].ToString();
                    t5.Text = MyData.dtmain.Rows[0]["c3"].ToString();
                }
                catch { }

                if (t5.Text == "")
                {
                    String name = "licschkre";
                    String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    "2.xlsx" +
                                    ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

                    OleDbConnection con = new OleDbConnection(constr);
                    OleDbCommand oconn = new OleDbCommand("Select Layihe, Adı From [" + name + "$] where Layihe Like '%" + t6.Text + "%'", con);
                    con.Open();
                    OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                    DataTable data = new DataTable();
                    data.Clear();
                    sda.Fill(data);
                    con.Close();

                    try
                    {
                        t6.Text = data.Rows[0]["Layihe"].ToString();
                        t5.Text = data.Rows[0]["Adı"].ToString();
                        //t5.Text = t5.Text.Substring(0, 1).ToUpper(ci) + t5.Text.Substring(1, t5.Text.Length - 1).ToLower(ci);
                    }
                    catch { }
                }

                if (t5.Text == "")
                {
                    String name2 = "Кредитный портфель";
                    String constr2 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    "1.xlsx" +
                                    ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

                    OleDbConnection con2 = new OleDbConnection(constr2);
                    OleDbCommand oconn2 = new OleDbCommand("Select [Наименование клиента], [Номер контракта] From [" + name2 + "$] WHERE [Номер контракта] Like '%" + t6.Text + "%'", con2);
                    con2.Open();
                    OleDbDataAdapter sda2 = new OleDbDataAdapter(oconn2);
                    DataTable data2 = new DataTable();
                    data2.Clear();
                    sda2.Fill(data2);
                    con2.Close();

                    try
                    {
                        t6.Text = data2.Rows[0]["Номер контракта"].ToString();
                        t5.Text = data2.Rows[0]["Наименование клиента"].ToString();
                    }
                    catch { }
                }
                try
                {
                    if (t5.Text.Substring(0, 1) == "'") t5.Text = t5.Text.Substring(1, t5.Text.Length - 1);
                    if (t6.Text.Substring(0, 1) == "'") t6.Text = t6.Text.Substring(1, t6.Text.Length - 1);
                }
                catch { }
            }
        }
        private void t4_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                t3.Text = "";
                MyData.selectCommand("baza.accdb", "SELECT * FROM etibarnameneqliyyat WHERE c4 like " + "'%" + t4.Text + "%'");

                MyData.dtmain = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmain);

                try
                {
                    t4.Text = MyData.dtmain.Rows[0]["c4"].ToString();
                    t3.Text = MyData.dtmain.Rows[0]["c3"].ToString();
                }
                catch { }

                if (t3.Text == "")
                {
                    String name = "licschkre";
                    String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    "2.xlsx" +
                                    ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

                    OleDbConnection con = new OleDbConnection(constr);
                    OleDbCommand oconn = new OleDbCommand("Select Layihe, Adı From [" + name + "$] where Layihe Like '%" + t4.Text + "%'", con);
                    con.Open();
                    OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                    DataTable data = new DataTable();
                    data.Clear();
                    sda.Fill(data);
                    con.Close();

                    try
                    {
                        t4.Text = data.Rows[0]["Layihe"].ToString();
                        t3.Text = data.Rows[0]["Adı"].ToString();
                        //t3.Text = t3.Text.Substring(0, 1).ToUpper(ci) + t3.Text.Substring(1, t3.Text.Length - 1).ToLower(ci);
                    }
                    catch { }
                }

                if (t3.Text == "")
                {
                    String name2 = "Кредитный портфель";
                    String constr2 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    "1.xlsx" +
                                    ";Extended Properties='Excel 12.0 xml;HDR=YES;';";

                    OleDbConnection con2 = new OleDbConnection(constr2);
                    OleDbCommand oconn2 = new OleDbCommand("Select [Наименование клиента], [Номер контракта] From [" + name2 + "$] WHERE [Номер контракта] Like '%" + t4.Text + "%'", con2);
                    con2.Open();
                    OleDbDataAdapter sda2 = new OleDbDataAdapter(oconn2);
                    DataTable data2 = new DataTable();
                    data2.Clear();
                    sda2.Fill(data2);
                    con2.Close();

                    try
                    {
                        t4.Text = data2.Rows[0]["Номер контракта"].ToString();
                        t3.Text = data2.Rows[0]["Наименование клиента"].ToString();
                    }
                    catch { }
                }
                try
                {
                    if (t3.Text.Substring(0, 1) == "'") t3.Text = t3.Text.Substring(1, t3.Text.Length - 1);
                    if (t4.Text.Substring(0, 1) == "'") t4.Text = t4.Text.Substring(1, t4.Text.Length - 1);
                }
                catch { }
            }
        }
        private void rbesli_CheckedChanged(object sender, EventArgs e)
        {
            txtqeydler.Text = "Sənədlərin əslini";
        }

        private void rbsureti_CheckedChanged(object sender, EventArgs e)
        {
            txtqeydler.Text = "Sənədlərin surətini";
        }

        private void cbTehvilAlanMutexessis_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbTehvilAlanMutexessis.Text == "Əsədov İbrahim Akif oğlu") cbvezifeTehvilAlan.Text = "Filial müdiri";
            else if (cbTehvilAlanMutexessis.Text == "Musayev Rəşad İslam oğlu") cbvezifeTehvilAlan.Text = "Baş menecer";
            else if(cbTehvilAlanMutexessis.Text == "Heydərov Namik Abisalam oğlu") cbvezifeTehvilAlan.Text = "Baş mütəxəssis";
            else if(cbTehvilAlanMutexessis.Text == "Şirinov Rasim Rafiqoviç") cbvezifeTehvilAlan.Text = "Sürücü";
            else if(cbTehvilAlanMutexessis.Text == "Elqar Ələkbərov") { cbvezifeTehvilAlan.Text = "Baş hüquqşünas"; cbsobeTehvilAlan2.Text = "Məhkəmə və İcra işləri Departamenti"; }
            else if(cbTehvilAlanMutexessis.Text == "Musayev Rəşad İslam oğlu") { cbvezifeTehvilAlan.Text = "Baş menecer"; cbsobeTehvilAlan2.Text = ""; }

        }

        private void btrefresh_Click(object sender, EventArgs e)
        {
            t1.Text = "";
            t2.Text = "";
            t3.Text = "";
            t4.Text = "";
            t5.Text = "";
            t6.Text = "";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string k = "", TVSobe = "", TASobe = "", TVqurum = "", TAqurum = "";
            if (rbesli.Checked == true) k = rbesli.Text;
            if (rbsureti.Checked == true) k = rbsureti.Text;

            DateTime dt = dttarix.Value.Date;
            string tarix = MyChange.TarixSozle(dt);

            try
            {
                TVqurum = cbsobeTehvilVeren2.Text;
                if (cbsobeTehvilVeren2.Text.Substring(0, 1) == "M") { TVqurum = "Hüquq"; }
                else if (cbsobeTehvilVeren2.Text.Substring(0, 1) == " ") { TVqurum = "Lizinq"; }
                else if (cbsobeTehvilVeren2.Text.Substring(0, 1) == "P") { TVqurum = "Problemi"; }
            }
            catch { }

            try
            {
                if (cbsobeTehvilAlan.Text.Substring(1, 6) == "AGBank") { TASobe = "“AGBank” ASC"; }
                else { TASobe = cbsobeTehvilAlan.Text; }
            }
            catch { }

            try
            {
                TAqurum = cbsobeTehvilAlan2.Text;
                if (cbsobeTehvilAlan2.Text.Substring(0, 1) == " ") { TAqurum = "Lizinq"; }
                else if (cbsobeTehvilAlan2.Text.Substring(0, 1) == "M") { TAqurum = "Hüquq"; }
                else if (cbsobeTehvilAlan2.Text.Substring(0, 1) == "P") { TAqurum = "Problemli"; }
            }
            catch { }

            if (t1.Text != "")
            {
                try
                {
                    MyData.insertCommand("baza.accdb", "insert into TehvilTeslim (a1, a2, a3, a4, a5, a6) Values ('" + t2.Text + "', '" + t1.Text + "', '" + dt.Day + " " + tarix + " " + dt.Year + "', '" + TVSobe + "/" + TVqurum + "/" + cbTehvilVerenMutexessis.Text + "','" + TASobe + "/" + TAqurum + "/" + cbTehvilAlanMutexessis.Text + "', '" + txtqeydler.Text + "')");
                }
                catch { MessageBox.Show("Tehvil Teslim bazaya qeyd olunarkən sehv..."); }

                try
                {
                    MyData.insertCommand("baza.accdb", "insert into Emeliyyatlar (a1, a2, a3) Values ('" + dt.Date + "', 'TƏHVİL TƏSLİM - " + t1.Text + " (" + t2.Text + ") - " + cbTehvilAlanMutexessis.Text + " (" + cbsobeTehvilAlan.Text + ") " + k + "','" + Environment.MachineName + "')");
                }
                catch { MessageBox.Show("Tehvil Teslim Emeliyyatlara qeyd olunarkən sehv..."); }
            }

            if (t3.Text != "")
            {
                try
                {
                    MyData.insertCommand("baza.accdb", "insert into TehvilTeslim (a1, a2, a3, a4, a5, a6) Values ('" + t4.Text + "', '" + t3.Text + "', '" + dt.Day + " " + tarix + " " + dt.Year + "', '" + TVSobe + "/" + TVqurum + "/" + cbTehvilVerenMutexessis.Text + "','" + TASobe + "/" + TAqurum + "/" + cbTehvilAlanMutexessis.Text + "', '" + txtqeydler.Text + "')");
                }
                catch { MessageBox.Show("Tehvil Teslim bazaya qeyd olunarkən sehv..."); }

                try
                {
                    MyData.insertCommand("baza.accdb", "insert into Emeliyyatlar (a1, a2, a3) Values ('" + dt.Date + "', 'TƏHVİL TƏSLİM - " + t3.Text + " (" + t4.Text + ") - " + cbTehvilAlanMutexessis.Text + " (" + cbsobeTehvilAlan.Text + ") " + k + "','" + Environment.MachineName + "')");
                }
                catch { MessageBox.Show("Tehvil Teslim Emeliyyatlara qeyd olunarkən sehv..."); }
            }

            if (t5.Text != "")
            {
                try
                {
                    MyData.insertCommand("baza.accdb", "insert into TehvilTeslim (a1, a2, a3, a4, a5, a6) Values ('" + t6.Text + "', '" + t5.Text + "', '" + dt.Day + " " + tarix + " " + dt.Year + "', '" + TVSobe + "/" + TVqurum + "/" + cbTehvilVerenMutexessis.Text + "','" + TASobe + "/" + TAqurum + "/" + cbTehvilAlanMutexessis.Text + "', '" + txtqeydler.Text + "')");
                }
                catch { MessageBox.Show("Tehvil Teslim bazaya qeyd olunarkən sehv..."); }

                try
                {
                    MyData.insertCommand("baza.accdb", "insert into Emeliyyatlar (a1, a2, a3) Values ('" + dt.Date + "', 'TƏHVİL TƏSLİM - " + t5.Text + " (" + t6.Text + ") - " + cbTehvilAlanMutexessis.Text + " (" + cbsobeTehvilAlan.Text + ") " + k + "','" + Environment.MachineName + "')");

                }
                catch { MessageBox.Show("Tehvil Teslim Emeliyyatlara qeyd olunarkən sehv..."); }
            }
        }

        private void Tehvil_Teslim_Load(object sender, EventArgs e)
        {

        }
    }
}