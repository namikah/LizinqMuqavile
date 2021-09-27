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
using System.Speech.Synthesis;
using System.Globalization;
using System.IO;
using Nsoft;

namespace Lizinq_Muqavile
{
    public partial class Kassa : Form
    {
        public Kassa()
        {
            InitializeComponent();
        }

        Excel.Application oXL;
        Excel._Workbook oWB;
        Excel._Worksheet oSheet;

        private void kassakitabi()  //---------------rasxodun nomresinin load olunmasi--------------
        {
            int i;
            double k=0, k2=0, k3=0;

            string commandText = "SELECT * FROM kassakitabi WHERE Tarix between #" + dtBaslama.Value.ToString("yyyy-MM-dd") + "# and #" + dtBitme.Value.AddDays(1).ToString("yyyy-MM-dd") + "#";
            commandText += " and Qeyd Like '%" + txtAxtar.Text + "%'";
            commandText += " order by Nomre desc";

            MyData.selectCommand("baza.accdb", commandText);
            MyData.dtmain = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmain);
            dataGridView1.DataSource = MyData.dtmain;

            for (i = 0; i < MyData.dtmain.Rows.Count; i++)
            {
                //kassanin mohkemlendirilmesi ucun edilen prixodlar hesablanmasin deye
                if (MyData.dtmain.Rows[i][2].ToString() != "AGBANK")
                {
                    k = k + Convert.ToDouble(MyData.dtmain.Rows[i][5]);
                    k2 = k2 + Convert.ToDouble(MyData.dtmain.Rows[i][3]);
                    k3 = k3 + Convert.ToDouble(MyData.dtmain.Rows[i][4]);
                }
            }

            label21.Text = Math.Round(Convert.ToDouble(k),2).ToString() + " AZN";
            label24.Text = Math.Round(Convert.ToDouble(k2),2).ToString()+ " AZN";
            label26.Text = Math.Round(Convert.ToDouble(k3),2).ToString() + " AZN";
            label28.Text = Math.Round(Convert.ToDouble(k - k2 - k3),2).ToString() + " AZN";
        
        }

        private void myrefresh()  //---------------rasxodun nomresinin load olunmasi--------------
        {
            string commanText = "Select * From kassanomre";

            MyData.selectCommand("baza.accdb", commanText);
            MyData.dtmain = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmain);

            try
            {
                txtnomre.Text = MyData.dtmain.Rows[0][0].ToString();
            }
            catch { MessageBox.Show("Məxaric nömrəsində səhv..."); };
        }

        private void QaliqRefresh()
        {
            string commanText = "Select * from kassaqaliq where AD='" + txtad.Text.ToString() + "'";

            MyData.selectCommand("baza.accdb", commanText);
            MyData.dtmain = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmain);

            try
            {
                txtqaliq.Text = Math.Round(Convert.ToDouble(MyData.dtmain.Rows[0][1]), 2).ToString();
            }
            catch { };
        }

        private void myrefresh2()
        {
            string commanText = "Select * from kassaqaliq where AD='" + cbsoyad.Text.ToString() + "'";

            MyData.selectCommand("baza.accdb", commanText);
            MyData.dtmain = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmain);

            try
            {
                txtkecmisavans.Text = MyData.dtmain.Rows[0][1].ToString();
            }
            catch { };

        }   // ---------------------myrefresh 2---------------kecmis avansin loadi--------------

        private void myrefresh3()
        {
            string commanText = "Select * From kassacek where cek like " + "'%" + textBox1.Text + "%'";

            MyData.selectCommand("baza.accdb", commanText);
            MyData.dtmain = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmain);

            try
            {
                textBox1.Text = MyData.dtmain.Rows[0][0].ToString();
            }
            catch { };
        }   //--------------------cekin load olunmasi-------------------------------------

        private void myrefresh4()
        {
            string commanText = "Select * From kassacek where cek like '%" + textBox2.Text + "%'";

            MyData.selectCommand("baza.accdb", commanText);
            MyData.dtmain = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmain);

            try
            {
                textBox2.Text = MyData.dtmain.Rows[0][0].ToString();
            }
            catch { };
        }   //--------------------cekin load olunmasi-------------------------------------

        private void myrefresh5()
        {
            string commanText = "Select * From kassacek where cek like '%" + textBox3.Text + "%'";
            MyData.selectCommand("baza.accdb", commanText);
            MyData.dtmain = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmain);

            try
            {
                textBox3.Text = MyData.dtmain.Rows[0][0].ToString();
            }
            catch { };
        }   //--------------------cekin load olunmasi-------------------------------------

        private void myrefresh6()
        {
            string commanText = "Select * From kassacek where cek like '%" + textBox4.Text + "%'";

            MyData.selectCommand("baza.accdb", commanText);
            MyData.dtmain = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmain);
            try
            {
                textBox4.Text = MyData.dtmain.Rows[0][0].ToString();
            }
            catch { };
        }   //--------------------cekin load olunmasi-------------------------------------

        private void myrefresh7()
        {
            string commanText = "Select * From kassacek where cek like '%" + textBox5.Text + "%'";

            MyData.selectCommand("baza.accdb", commanText);
            MyData.dtmain = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmain);

            try
            {
                textBox5.Text = MyData.dtmain.Rows[0][0].ToString();
            }
            catch { };
        }   //--------------------cekin load olunmasi-------------------------------------

        private void myrefresh8()
        {
            string commanText = " Select * From AGLizinqKassa";

            MyData.selectCommand("baza.accdb", commanText);
            MyData.dtmain = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmain);

            try
            {
                button19.Text = MyData.dtmain.Rows[0][0].ToString();
                label32.Text = "AGLizinq QSC Kassada qalıq - " + button19.Text + " AZN" ;
            }
            catch { };

            
        }   //--------------------cekin load olunmasi-------------------------------------

        private void kassa1_Load(object sender, EventArgs e)
        {
            myrefresh();
            kassakitabi();
            myrefresh8();
            myrefresh2();
            QaliqRefresh();

            /*
            DateTime dt = DateTime.Now;
            dttarix.Text = dt.Date.ToShortDateString();
            dttarixhesabat.Text = dt.Date.ToShortDateString();
            dttarixavansalinma.Text = dt.Date.ToShortDateString();
            dateTimePicker1.Text = dt.Date.ToShortDateString();
            dateTimePicker2.Text = dt.Date.ToShortDateString();
            dateTimePicker3.Text = dt.Date.ToShortDateString();
            dateTimePicker4.Text = dt.Date.ToShortDateString();
            dateTimePicker5.Text = dt.Date.ToShortDateString();*/

        }

        private void reqemler()      //------reqem yazi ile---------------------------------------------------------------
        {
            txt2.Text = MyChange.ReqemToMetn(Convert.ToDouble(txt1.Text));
        }

        private void Print()    //-----------------Print rasxod ucun-------------------------------
        {
            try
            {
                File.Copy("Rasxod.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Məxaric - " + txtad.Text + ".xlsx", true);
            }
            catch { MessageBox.Show("Rasxod.xlsx tapılmadı."); }

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Məxaric - " + txtad.Text + ".xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

            if (txt1.Text == "") txt1.Text = "0";
            if (txt2.Text == "") txt2.Text = "Sıfır";
            if (txtad.Text == "") txtad.Text = "Şirinov Rasim Rafiqoviç";
            oSheet.Cells[4, 7] = txtnomre.Text;
            oSheet.Cells[7, 6] = dttarix.Text;
            oSheet.Cells[28, 4] = dttarix.Text;
            oSheet.Cells[12, 4] = txtad.Text.ToString();
            oSheet.Cells[10, 6] = txt1.Text.ToString();
            oSheet.Cells[17, 3] = txt2.Text.ToString();
            oSheet.Cells[25, 3] = txt2.Text.ToString();
            oSheet.Cells[15, 4] = txtesasname.Text.ToString();

        //----------------------------------------------------------------excelde sutunlarin silinmesi 
            Excel.Range range = oSheet.get_Range("A37:AB37", Type.Missing);
            range.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
       
            if (txtad.Text == "Əsədov İbrahim Akif oğlu")
                oSheet.Cells[30, 3] = "Səxsiyyət Vəsiqəsi AZE № 13220118, BAKI-Nizami  RPI 27.05.2013-cü il";
            else if (txtad.Text == "Musayev Rəşad İslam oğlu")
                oSheet.Cells[30, 3] = "Səxsiyyət Vəsiqəsi AZE № 17269313 ASAN3 29.09.2017";
            else if (txtad.Text == "Rəşidov Zaur Xanəliyeviç")
                oSheet.Cells[30, 3] = "Səxsiyyət Vəsiqəsi AZE № 08272442 Səbail RPI 06.04.2011";
            else if (txtad.Text == "Heydərov Namik Abisalam oğlu")
                oSheet.Cells[30, 3] = "Səxsiyyət Vəsiqəsi AZE № 13726303 Yasamal RPI 29.03.2013";
            else if (txtad.Text == "Şirinov Rasim Rafiqoviç")
                oSheet.Cells[30, 3] = "Səxsiyyət Vəsiqəsi AZE № 08203665, BAKI-XƏTAI  RPI 16.12.2010-cu il";
            else if (txtad.Text == "Səfərli Nərmin Səfər qızı")
                oSheet.Cells[30, 3] = "Səxsiyyət Vəsiqəsi AZE № 09513223, BAKI-XƏTAI  RPI 29.08.2012-ci il";
            else if (txtad.Text == "Məmmədov İlqar İmran oğlu")
                oSheet.Cells[30, 3] = "Səxsiyyət Vəsiqəsi AZE № 09229670, SUMQAYIT ŞPİ QŞVŞ 02.09.2011-ci il";


            string s1;
            int a;
            s1 = "'" + txtnomre.Text + "'";
            a = Convert.ToInt32(txtnomre.Text);
            a = a + 1;

         //   oSheet.PrintOut();
         //   oWB.Close(SaveChanges: false);
         //   oXL.Workbooks.Close();
            try
            {
                MyData.updateCommand("baza.accdb", "UPDATE kassanomre  SET a1 ='" + a.ToString() + "'");
            }
            catch { };

            MyData.updateCommand("baza.accdb", "UPDATE kassaqaliq  SET Qaliq =" + "'" + Math.Round((Convert.ToDouble(txtqaliq.Text) + Convert.ToDouble(txt1.Text)), 2).ToString() + "'" + " WHERE AD like " + "'" + txtad.Text.ToString() + "'");

            ///AGLizinq qaliq ucun-----------------------
            MyData.updateCommand("baza.accdb", "UPDATE AGLizinqKassa  SET a1 =" + "'" + Math.Round((Convert.ToDouble(button19.Text) - Convert.ToDouble(txt1.Text)), 2).ToString() + "'");
            myrefresh8();
            myrefresh();
        }

        private void Printsade()    //-----------------Print excel ucun-------------------------------
        {
            try
            {
                File.Copy("Rasxod.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Məxaric - " + txtad.Text + ".xlsx", true);
            }
            catch { MessageBox.Show("Rasxod.xlsx tapılmadı."); }

                //Get a new workbook.
                oXL = new Excel.Application();
                oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Məxaric - " + txtad.Text + ".xlsx"));
                oSheet = (Excel._Worksheet)oWB.Sheets[1];
                oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                if (txt1.Text == "") txt1.Text = "0";
                if (txt2.Text == "") txt2.Text = "Sıfır";
                if (txtad.Text == "") txtad.Text = "Şirinov Rasim Rafiqoviç";
                oSheet.Cells[4, 7] = txtnomre.Text;
                oSheet.Cells[7, 6] = dttarix.Text;
                oSheet.Cells[28, 4] = dttarix.Text;
                oSheet.Cells[12, 4] = txtad.Text.ToString();
                oSheet.Cells[10, 6] = txt1.Text.ToString();
                oSheet.Cells[17, 3] = txt2.Text.ToString();
                oSheet.Cells[25, 3] = txt2.Text.ToString();
                oSheet.Cells[15, 4] = txtesasname.Text.ToString();

                if (txtad.Text == "Əsədov İbrahim Akif oğlu")
                    oSheet.Cells[30, 3] = "Səxsiyyət Vəsiqəsi AZE № 13220118, BAKI-Nizami  RPI 27.05.2013-cü il";
            else if (txtad.Text == "Musayev Rəşad İslam oğlu")
                    oSheet.Cells[30, 3] = "Səxsiyyət Vəsiqəsi AZE № 17269313 ASAN3 29.09.2017";
            else if (txtad.Text == "Rəşidov Zaur Xanəliyeviç")
                    oSheet.Cells[30, 3] = "Səxsiyyət Vəsiqəsi AZE № 08272442 Səbail RPI 06.04.2011";
            else if (txtad.Text == "Heydərov Namik Abisalam oğlu")
                    oSheet.Cells[30, 3] = "Səxsiyyət Vəsiqəsi AZE № 13726303 Yasamal RPI 29.03.2013";
            else if (txtad.Text == "Şirinov Rasim Rafiqoviç")
                    oSheet.Cells[30, 3] = "Səxsiyyət Vəsiqəsi AZE № 08203665, BAKI-XƏTAI  RPI 16.12.2010-cu il";
            else if (txtad.Text == "Səfərli Nərmin Səfər qızı")
                    oSheet.Cells[30, 3] = "Səxsiyyət Vəsiqəsi AZE № 09513223, BAKI-XƏTAI  RPI 29.08.2012-ci il";
            else if (txtad.Text == "Məmmədov İlqar İmran oğlu")
                    oSheet.Cells[30, 3] = "Səxsiyyət Vəsiqəsi AZE № 09229670, SUMQAYIT ŞPİ QŞVŞ 02.09.2011-ci il";



                string s1;
                int a;
                s1 = "'" + txtnomre.Text + "'";
                a = Convert.ToInt32(txtnomre.Text);
                a = a + 1;


                oXL.Visible = false;
                oSheet.Activate();
                oSheet.Range["A1"].Select();
                oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
                oSheet.PrintOut();
                try
                {
                    oXL.DisplayAlerts = false;
                    oWB.Close(SaveChanges: true);
                    oXL.Application.Quit(); 
                    oWB.Save();
                }
                catch { }




                try
                {
                MyData.updateCommand("baza.accdb", "UPDATE kassanomre  SET a1 ='" + a.ToString() + "'");
                }
                catch { };
            MyData.updateCommand("baza.accdb", "UPDATE kassaqaliq  SET Qaliq =" + "'" + Math.Round((Convert.ToDouble(txtqaliq.Text) + Convert.ToDouble(txt1.Text)), 2).ToString() + "'" + " WHERE AD like " + "'" + txtad.Text.ToString() + "'");

            MyData.updateCommand("baza.accdb", "UPDATE AGLizinqKassa  SET a1 =" + "'" + Math.Round((Convert.ToDouble(button19.Text) - Convert.ToDouble(txt1.Text)), 2).ToString() + "'");
                myrefresh8();
                ///AGLizinq qaliq ucun-----------------------
                myrefresh();
            
        }

        private void Print2()    //-----------------Print Avans ucun-------------------------------
        {
            try
            {
                File.Copy("Avans.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Avans - " + cbsoyad.Text + ".xlsx", true);
            }
            catch { MessageBox.Show("Avans.xlsx tapılmadı."); }

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Avans - " + cbsoyad.Text + ".xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

            if (txtavansalinib.Text == "") txtavansalinib.Text = "0";
            if (txtavansqaytarilib.Text == "") txtavansqaytarilib.Text = "0";
            if (txtkecmisavans.Text == "") txtkecmisavans.Text = "0";
            if (txtavansxerclenib.Text == "") txtavansxerclenib.Text = "0";

            oSheet.Cells[10, 1] = "AVANS HESABATI №  " + dttarixhesabat.Text + "-cu il.";
            if (txtavansalinib.Text != "0") oSheet.Cells[17, 2] = "( " + dttarixavansalinma.Text + " )"; else oSheet.Cells[17, 2] = "()";
            oSheet.Cells[8, 3] = cbsoyad.Text.ToString();
            oSheet.Cells[5, 6] = cbvezife.Text.ToString();
            oSheet.Cells[14, 5] = (Convert.ToDouble(txtkecmisavans.Text) - Convert.ToDouble(txtavansalinib.Text)).ToString();
            oSheet.Cells[17, 5] = txtavansalinib.Text.ToString();
            oSheet.Cells[26, 5] = txtavansxerclenib.Text.ToString();
            oSheet.Cells[27, 5] = txtavansqaytarilib.Text.ToString();

            if (textBox1.Text != "")
            {
                oSheet.Cells[41, 1] = dateTimePicker1.Text.ToString();
                oSheet.Cells[41, 3] = textBox1.Text.ToString();
                oSheet.Cells[41, 11] = txtm1.Text.ToString();
            }
            if (textBox2.Text != "")
            {
                oSheet.Cells[42, 1] = dateTimePicker2.Text.ToString();
                oSheet.Cells[42, 3] = textBox2.Text.ToString();
                oSheet.Cells[42, 11] = txtm2.Text.ToString();
            }
            if (textBox3.Text != "")
            {
                oSheet.Cells[43, 1] = dateTimePicker3.Text.ToString();
                oSheet.Cells[43, 3] = textBox3.Text.ToString();
                oSheet.Cells[43, 11] = txtm3.Text.ToString();
            }
            if (textBox4.Text != "")
            {
                oSheet.Cells[44, 1] = dateTimePicker4.Text.ToString();
                oSheet.Cells[44, 3] = textBox4.Text.ToString();
                oSheet.Cells[44, 11] = txtm4.Text.ToString();
            }
            if (textBox5.Text != "")
            {
                oSheet.Cells[45, 1] = dateTimePicker5.Text.ToString();
                oSheet.Cells[45, 3] = textBox5.Text.ToString();
                oSheet.Cells[45, 11] = txtm5.Text.ToString();
            }


            //  oSheet.PrintOut();
            //   oWB.Close(SaveChanges: false);
            //   oXL.Workbooks.Close();
            MyData.updateCommand("baza.accdb", "UPDATE kassaqaliq  SET Qaliq = " + "'" + Math.Round((Convert.ToDouble(txtkecmisavans.Text) - Convert.ToDouble(txtavansxerclenib.Text) - Convert.ToDouble(txtavansqaytarilib.Text)), 2).ToString() + "'" + " WHERE AD like " + "'" + cbsoyad.Text.ToString() + "'");

            ///AGLizinq qaliq ucun-----------------------
            if (txtavansqaytarilib.Text != "0" && txtavansqaytarilib.Text != "")
            {
                MyData.updateCommand("baza.accdb", "UPDATE AGLizinqKassa  SET a1 =" + "'" + Math.Round((Convert.ToDouble(button19.Text) + Convert.ToDouble(txtavansqaytarilib.Text)), 2).ToString() + "'");
                myrefresh8();
            }
            ///AGLizinq qaliq ucun-----------------------

            try
            {
                if (textBox1.Text != "")
                {
                    MyData.selectCommand("baza.accdb", " Select * From kassacek where cek Like '%" + textBox1.Text.Substring(0, 10).ToString() + "%'");
                    MyData.dtmain = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmain);

                    if (MyData.dtmain.Rows.Count > 0)
                    {
                        try
                        {
                            MyData.updateCommand("baza.accdb", "UPDATE kassacek SET cek = '" + textBox1.Text + "' where cek Like '%" + textBox1.Text.Substring(0, 10).ToString() + "%'");
                        }
                        catch { }
                    }
                    else
                    {
                        try
                        {
                            MyData.insertCommand("baza.accdb", "insert into kassacek (cek) Values ('" + textBox1.Text + "')");
                            
                            
                        }
                        catch { }
                    }

                }
            }
            catch { }

            try
            {
                if (textBox2.Text != "")
                {
                    MyData.selectCommand("baza.accdb", " Select * From kassacek where cek like " + "'%" + textBox2.Text.Substring(0, 10).ToString() + "%'");
                    MyData.dtmain = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmain);

                    if (MyData.dtmain.Rows.Count > 0)
                    {
                        try
                        {
                            MyData.updateCommand("baza.accdb", "UPDATE kassacek SET cek = '" + textBox2.Text + "' where cek Like '%" + textBox2.Text.Substring(0, 10).ToString() + "%'");
                        }
                        catch { }
                    }
                    else
                    {
                        try
                        {
                            MyData.insertCommand("baza.accdb", "insert into kassacek (cek) Values ('" + textBox2.Text + "')");
                            
                            
                        }
                        catch { }
                    }

                }
            }
            catch { }

            try
            {
                if (textBox3.Text != "")
                {
                    MyData.selectCommand("baza.accdb", " Select * From kassacek where cek like " + "'%" + textBox3.Text.Substring(0, 10).ToString() + "%'");
                    MyData.dtmain = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmain);

                    if (MyData.dtmain.Rows.Count > 0)
                    {
                        try
                        {
                            MyData.updateCommand("baza.accdb", "UPDATE kassacek SET cek = '" + textBox3.Text + "' where cek Like '%" + textBox3.Text.Substring(0, 10).ToString() + "%'");
                        }
                        catch { }
                    }
                    else
                    {
                        try
                        {
                            MyData.insertCommand("baza.accdb", "insert into kassacek (cek) Values ('" + textBox3.Text + "')");
                            
                        }
                        catch { }
                    }
                }
            }
            catch { }

            try
            {
                if (textBox4.Text != "")
                {
                    MyData.selectCommand("baza.accdb", " Select * From kassacek where cek like " + "'%" + textBox4.Text.Substring(0, 10).ToString() + "%'");
                    MyData.dtmain = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmain);

                    if (MyData.dtmain.Rows.Count > 0)
                    {
                        try
                        {
                            MyData.updateCommand("baza.accdb", "UPDATE kassacek SET cek = '" + textBox4.Text + "' where cek Like '%" + textBox4.Text.Substring(0, 10).ToString() + "%'");
                        }
                        catch { }
                    }
                    else
                    {
                        try
                        {
                            MyData.insertCommand("baza.accdb", "insert into kassacek (cek) Values ('" + textBox4.Text + "')");
                            
                            
                        }
                        catch { }
                    }
                }
            }
            catch { }

            try
            {
                if (textBox5.Text != "")
                {
                    MyData.selectCommand("baza.accdb", " Select * From kassacek where cek like " + "'%" + textBox5.Text.Substring(0, 10).ToString() + "%'");
                    MyData.dtmain = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmain);

                    if (MyData.dtmain.Rows.Count > 0)
                    {
                        try
                        {
                            MyData.updateCommand("baza.accdb", "UPDATE kassacek SET cek = '" + textBox5.Text + "' where cek Like '%" + textBox5.Text.Substring(0, 10).ToString() + "%'");
                        }
                        catch { }
                    }
                    else
                    {
                        try
                        {
                            MyData.insertCommand("baza.accdb", "insert into kassacek (cek) Values ('" + textBox5.Text + "')");
                            
                            
                        }
                        catch { }
                    }
                }
            }
            catch { }
            myrefresh2();
        }

        private void Print2sade()    //-----------------Print Avans pramoy print ucun-------------------------------
        {
            try
            {
                File.Copy("Avans.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Avans - " + cbsoyad.Text + ".xlsx", true);
            }
            catch { MessageBox.Show("Avans.xlsx tapılmadı."); }

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Avans - " + cbsoyad.Text + ".xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            if (txtavansalinib.Text == "") txtavansalinib.Text = "0";
            if (txtavansqaytarilib.Text == "") txtavansqaytarilib.Text = "0";
            if (txtkecmisavans.Text == "") txtkecmisavans.Text = "0";
            if (txtavansxerclenib.Text == "") txtavansxerclenib.Text = "0";

            oSheet.Cells[10, 1] = "AVANS HESABATI №  " + dttarixhesabat.Text + "-cu il.";
            if (txtavansalinib.Text != "0") oSheet.Cells[17, 2] = "( " + dttarixavansalinma.Text + " )"; else oSheet.Cells[17, 2] = "()";
            oSheet.Cells[8, 3] = cbsoyad.Text.ToString();
            oSheet.Cells[5, 6] = cbvezife.Text.ToString();
            oSheet.Cells[14, 5] = (Convert.ToDouble(txtkecmisavans.Text) - Convert.ToDouble(txtavansalinib.Text)).ToString();
            oSheet.Cells[17, 5] = txtavansalinib.Text.ToString();
            oSheet.Cells[26, 5] = txtavansxerclenib.Text.ToString();
            oSheet.Cells[27, 5] = txtavansqaytarilib.Text.ToString();

            if (textBox1.Text != "")
            {
                oSheet.Cells[41, 1] = dateTimePicker1.Text.ToString();
                oSheet.Cells[41, 3] = textBox1.Text.ToString();
                oSheet.Cells[41, 11] = txtm1.Text.ToString();
            }
            if (textBox2.Text != "")
            {
                oSheet.Cells[42, 1] = dateTimePicker2.Text.ToString();
                oSheet.Cells[42, 3] = textBox2.Text.ToString();
                oSheet.Cells[42, 11] = txtm2.Text.ToString();
            }
            if (textBox3.Text != "")
            {
                oSheet.Cells[43, 1] = dateTimePicker3.Text.ToString();
                oSheet.Cells[43, 3] = textBox3.Text.ToString();
                oSheet.Cells[43, 11] = txtm3.Text.ToString();
            }
            if (textBox4.Text != "")
            {
                oSheet.Cells[44, 1] = dateTimePicker4.Text.ToString();
                oSheet.Cells[44, 3] = textBox4.Text.ToString();
                oSheet.Cells[44, 11] = txtm4.Text.ToString();
            }
            if (textBox5.Text != "")
            {
                oSheet.Cells[45, 1] = dateTimePicker5.Text.ToString();
                oSheet.Cells[45, 3] = textBox5.Text.ToString();
                oSheet.Cells[45, 11] = txtm5.Text.ToString();
            }

            oXL.Visible = false;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
            oSheet.PrintOut();

            try
            {
                oXL.DisplayAlerts = false;
                oWB.Close(SaveChanges: true);
                oXL.Application.Quit();
                oWB.Save();
            }
            catch { }
            MyData.updateCommand("baza.accdb", "UPDATE kassaqaliq  SET Qaliq = " + "'" + Math.Round((Convert.ToDouble(txtkecmisavans.Text) - Convert.ToDouble(txtavansxerclenib.Text) - Convert.ToDouble(txtavansqaytarilib.Text)), 2).ToString() + "'" + " WHERE AD like " + "'" + cbsoyad.Text.ToString() + "'");


            ///AGLizinq qaliq ucun-----------------------
            if (txtavansqaytarilib.Text != "0" && txtavansqaytarilib.Text != "")
            {
                MyData.updateCommand("baza.accdb", "UPDATE AGLizinqKassa  SET a1 =" + "'" + Math.Round((Convert.ToDouble(button19.Text) + Convert.ToDouble(txtavansqaytarilib.Text)), 2).ToString() + "'");
                myrefresh8();
            }
            ///AGLizinq qaliq ucun-----------------------


            if (textBox1.Text != "")
            {
                MyData.selectCommand("baza.accdb", " Select * From kassacek where cek Like '%" + textBox1.Text.Substring(0, 10).ToString() + "%'");
                MyData.dtmain = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmain);

                if (MyData.dtmain.Rows.Count > 0)
                {
                    try
                    {
                        MyData.updateCommand("baza.accdb", "UPDATE kassacek SET cek = '" + textBox1.Text + "' where cek Like '%" + textBox1.Text.Substring(0, 10).ToString() + "%'");
                    }
                    catch { }
                }
                else
                {
                    try
                    {
                        MyData.insertCommand("baza.accdb", "insert into kassacek (cek) Values ('" + textBox1.Text + "')");
                        
                        
                    }
                    catch { }
                }

            }
            if (textBox2.Text != "")
            {



                MyData.selectCommand("baza.accdb", " Select * From kassacek where cek like " + "'%" + textBox2.Text.Substring(0, 10).ToString() + "%'");
                MyData.dtmain = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmain);

                if (MyData.dtmain.Rows.Count > 0)
                {
                    try
                    {
                        MyData.updateCommand("baza.accdb", "UPDATE kassacek SET cek = '" + textBox2.Text + "' where cek Like '%" + textBox2.Text.Substring(0, 10).ToString() + "%'");
                        
                        
                    }
                    catch { }
                }
                else
                {
                    try
                    {
                        MyData.insertCommand("baza.accdb", "insert into kassacek (cek) Values ('" + textBox2.Text + "')");
                        
                        
                    }
                    catch { }
                }

            }
            if (textBox3.Text != "")
            {


                MyData.selectCommand("baza.accdb", " Select * From kassacek where cek like " + "'%" + textBox3.Text.Substring(0, 10).ToString() + "%'");
                MyData.dtmain = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmain);

                if (MyData.dtmain.Rows.Count > 0)
                {
                    try
                    {
                        MyData.updateCommand("baza.accdb", "UPDATE kassacek SET cek = '" + textBox3.Text + "' where cek Like '%" + textBox3.Text.Substring(0, 10).ToString() + "%'");
                        
                        
                    }
                    catch { }
                }
                else
                {
                    try
                    {
                        MyData.insertCommand("baza.accdb", "insert into kassacek (cek) Values ('" + textBox3.Text + "')");
                    }
                    catch { }
                }
            }
            if (textBox4.Text != "")
            {
                MyData.selectCommand("baza.accdb", " Select * From kassacek where cek like " + "'%" + textBox4.Text.Substring(0, 10).ToString() + "%'");
                MyData.dtmain = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmain);

                if (MyData.dtmain.Rows.Count > 0)
                {
                    try
                    {
                        MyData.updateCommand("baza.accdb", "UPDATE kassacek SET cek = '" + textBox4.Text + "' where cek Like '%" + textBox4.Text.Substring(0, 10).ToString() + "%'");
                        
                        
                    }
                    catch { }
                }
                else
                {
                    try
                    {
                        MyData.insertCommand("baza.accdb", "insert into kassacek (cek) Values ('" + textBox4.Text + "')");
                        
                        
                    }
                    catch { }
                }
            }
            if (textBox5.Text != "")
            {
                MyData.selectCommand("baza.accdb", " Select * From kassacek where cek like " + "'%" + textBox5.Text.Substring(0, 10).ToString() + "%'");
                MyData.dtmain = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmain);

                if (MyData.dtmain.Rows.Count > 0)
                {
                    try
                    {
                        MyData.updateCommand("baza.accdb", "UPDATE kassacek SET cek = '" + textBox5.Text + "' where cek Like '%" + textBox5.Text.Substring(0, 10).ToString() + "%'");
                    }
                    catch { }
                }
                else
                {
                    try
                    {
                        MyData.insertCommand("baza.accdb", "insert into kassacek (cek) Values ('" + textBox5.Text + "')"); 
                    }
                    catch { }
                }
            }

            myrefresh2();
        }

        private void btprint_Click(object sender, EventArgs e)
        {
            if (txtad.Text == "") { MessageBox.Show(" Soyadı A.A  Qeyd olunmayıb ... "); return; }

            DialogResult result2 = MessageBox.Show("Davam etmək istəyirsiniz?", "?", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (result2 == DialogResult.No) { return; }

                Printsade();

                //--------------------qaliq refres olunsun rasxod sehifede-----------------------------------------------
                cbsoyad.Text = txtad.Text;
                txtad.ForeColor = Color.DarkRed;

                MyData.selectCommand("baza.accdb", "Select * from kassaqaliq where AD = " + "'" + txtad.Text.ToString() + "'");
                MyData.dtmain = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmain);
                try
                {
                    txtqaliq.Text = Math.Round(Convert.ToDouble(MyData.dtmain.Rows[0][1]),2).ToString();
                }
                catch { };
                txtkecmisavans.Text = MyData.dtmain.Rows[0][1].ToString();

                MyData.insertCommand("baza.accdb", "insert into kassakitabi (Tarix, İstifadəçi, Mədaxil, Xərclənib, Məxaric, Qalıq, Qeyd) Values ('" + dttarix.Text + "'," + "'" + txtad.Text + "'," + "'" + "0" + "'," + "'" + "0" + "'," + "'" + Math.Round(Convert.ToDouble(txt1.Text), 2).ToString() + "'," + "'" + txtqaliq.Text + "','Kassa Qalıq - " + button19.Text + " AZN / " + richTextBox1.Text + "')");

                MyData.insertCommand("baza.accdb", "insert into Emeliyyatlar (a1, a2, a3) Values ('" + DateTime.Now + "', 'MƏXARİC - " + txtad.Text + " - " + Math.Round(Convert.ToDouble(txt1.Text), 2).ToString() + " manat ( " + richTextBox1.Text + " )','" + Environment.MachineName + "')");
                
                kassakitabi();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            switch (txtad.Text)
            {
                case "Əsədov İbrahim Akif oğlu": pictureBox1.ImageLocation = "Images\\5.jpg"; break;
                case "Musayev Rəşad İslam oğlu": pictureBox1.ImageLocation = "Images\\2.jpg"; break;
                case "Rəşidov Zaur Xanəliyeviç": pictureBox1.ImageLocation = "Images\\4.jpg"; break;
                case "Heydərov Namik Abisalam oğlu": pictureBox1.ImageLocation = "Images\\3.jpg"; break;
                case "Şirinov Rasim Rafiqoviç": pictureBox1.ImageLocation = "Images\\1.jpg"; break;
                case "Səfərli Nərmin Səfər qızı": pictureBox1.ImageLocation = "Images\\6.jpg"; break;
                case "Məmmədov İlqar İmran oğlu": pictureBox1.ImageLocation = "Images\\7.jpg"; break;
                default: break;
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if (dttarixavansalinma.Enabled == false) dttarixavansalinma.Enabled = true; else dttarixavansalinma.Enabled = false;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (dttarix.Enabled == false) dttarix.Enabled = true; else dttarix.Enabled = false;
        }

        private void button13_Click(object sender, EventArgs e)
        {
            if (txtad.Enabled == false) txtad.Enabled = true; else txtad.Enabled = false;
          
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (txtnomre.Enabled == false) txtnomre.Enabled = true; else txtnomre.Enabled = false;
        }

        private void txt1_TextChanged(object sender, EventArgs e)
        {
            reqemler();
            txtavansalinib.Text = txt1.Text;
            txtavansxerclenib.Text = txt1.Text;
        }

        private void txtad_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch(txtad.Text)
            {
                case "Əsədov İbrahim Akif oğlu": pictureBox1.ImageLocation = "Images\\5.jpg"; break;
                case "Musayev Rəşad İslam oğlu": pictureBox1.ImageLocation = "Images\\2.jpg"; break;
                case "Rəşidov Zaur Xanəliyeviç": pictureBox1.ImageLocation = "Images\\4.jpg"; break;
                case "Heydərov Namik Abisalam oğlu": pictureBox1.ImageLocation = "Images\\3.jpg"; break;
                case "Şirinov Rasim Rafiqoviç": pictureBox1.ImageLocation = "Images\\1.jpg"; break;
                case "Səfərli Nərmin Səfər qızı": pictureBox1.ImageLocation = "Images\\6.jpg"; break;
                case "Məmmədov İlqar İmran oğlu": pictureBox1.ImageLocation = "Images\\7.jpg"; break;
                default: break;
            }

            cbsoyad.Text = txtad.Text;
            txtad.ForeColor = Color.DarkRed;
            MyData.selectCommand("baza.accdb", "Select * from kassaqaliq where AD = '" + txtad.Text.ToString() + "'");
            MyData.dtmain = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmain);
            try
            {
                txtqaliq.Text = Math.Round(Convert.ToDouble(MyData.dtmain.Rows[0][1]),2).ToString();
            }
            catch { };
        }

        private void txtad_MouseHover(object sender, EventArgs e)
        {
            txtad.ForeColor = Color.DarkGreen;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (dttarixhesabat.Enabled == false) dttarixhesabat.Enabled = true; else dttarixhesabat.Enabled = false;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (txtkecmisavans.Enabled == false) txtkecmisavans.Enabled = true; else txtkecmisavans.Enabled = false;
        }

        private void cbsoyad_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtad.Text = cbsoyad.Text;
            cbsoyad.ForeColor = Color.DarkGreen;

            if (cbsoyad.Text == "Əsədov İbrahim Akif oğlu")
                cbvezife.Text = "Filial müdiri";
            else if (cbsoyad.Text == "Musayev Rəşad İslam oğlu")
                cbvezife.Text = "Baş menecer";
            else if (cbsoyad.Text == "Rəşidov Zaur Xanəliyeviç")
                cbvezife.Text = "Aparıcı lizinq mutəxəssisi";
            else if (cbsoyad.Text == "Heydərov Namik Abisalam oğlu")
                cbvezife.Text = "Baş lizinq mutəxəssisi";
            else if (cbsoyad.Text == "Şirinov Rasim Rafiqoviç")
                cbvezife.Text = "Sürücü";
            else if (cbsoyad.Text == "Səfərli Nərmin Səfər qızı")
                cbvezife.Text = "Kiçik lizinq mütəxəssisi";
            else if (cbsoyad.Text == "Məmmədov İlqar İmran oğlu")
                cbvezife.Text = "";

            myrefresh2();
        }

        private void cbvezife_SelectedIndexChanged(object sender, EventArgs e)
        {
            cbvezife.ForeColor = Color.DarkGreen;
        }

        private void cbsoyad_MouseHover(object sender, EventArgs e)
        {
            cbsoyad.ForeColor = Color.DarkGreen;
        }

        private void cbvezife_MouseHover(object sender, EventArgs e)
        {
            cbvezife.ForeColor = Color.DarkGreen;
        }

        private void btprint2_Click(object sender, EventArgs e)
        {
            if (cbsoyad.Text == "") { MessageBox.Show(" Soyadı A.A  Qeyd olunmayıb ... "); return; }

            DialogResult result2 = MessageBox.Show("Davam etmək istəyirsiniz?", "?", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (result2 == DialogResult.No) { return; }

            Print2sade();

            MyData.selectCommand("baza.accdb", "Select * from kassaqaliq where AD = " + "'" + txtad.Text.ToString() + "'");
            MyData.dtmain = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmain);
            try
            {
                txtqaliq.Text = Math.Round(Convert.ToDouble(MyData.dtmain.Rows[0][1]), 2).ToString();
            }
            catch { };

            MyData.insertCommand("baza.accdb", "insert into kassakitabi (Tarix, İstifadəçi, Mədaxil, Xərclənib, Məxaric, Qalıq, Qeyd) Values ('" + dttarixhesabat.Text + "','" + cbsoyad.Text + "','" + txtavansqaytarilib.Text + "','" + txtavansxerclenib.Text + "','" + "0" + "'," + "'" + txtqaliq.Text + "','Kassa Qalıq - " + button19.Text + " AZN / Verilmiş avansın bağlanması')");

            MyData.selectCommand("baza.accdb", "insert into Emeliyyatlar (a1, a2, a3) Values ('" + DateTime.Now + "', 'AVANS - " + cbsoyad.Text + " - xərclənib " + txtavansxerclenib.Text + " manat - qaytarılıb " + txtavansqaytarilib.Text + " manat','" + Environment.MachineName + "')");

            kassakitabi();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (cbvezife.Enabled == false) cbvezife.Enabled = true; else cbvezife.Enabled = false;
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) myrefresh3();
        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) myrefresh4();
        }

        private void textBox3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) myrefresh5();
        }

        private void textBox4_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) myrefresh6();
        }

        private void textBox5_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) myrefresh7();
        }

        private void txtm1_MouseClick(object sender, MouseEventArgs e)
        {
            try
            {
                txtm1.Text = txtavansxerclenib.Text;
                if (txtavansxerclenib.Text == "0") txtm1.Text = "";
                if (txtavansqaytarilib.Text == "") { MessageBox.Show("Qaytarılan avans məbləği düzgün deyil..."); txtavansqaytarilib.Text = "0"; }
                if (txtavansxerclenib.Text == "") { MessageBox.Show("Xərclənən avans məbləği düzgün deyil..."); txtavansxerclenib.Text = "0"; }
            }
            catch { };
        }
        private void button10_Click(object sender, EventArgs e)
        {
            if (txtesasname.Enabled == false) txtesasname.Enabled = true; else txtesasname.Enabled = false;
        }

        private void haqqındaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("'AQLizinq' QSC-nin kassa Əməliyyatları..");
        }

        private void helpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Press F1");
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            if (pictureBox1.Left != 190)
            {
                pictureBox1.Top = 3;
                pictureBox1.Left = 190;
                pictureBox1.Width = 453;
                pictureBox1.Height = 247;
                timer2.Enabled = false;
                return;
            }
            if (pictureBox1.Left != 526)
            {
                pictureBox1.Top = 3;
                pictureBox1.Left = 526;
                pictureBox1.Width = 117;
                pictureBox1.Height = 72;
                timer2.Enabled = false;

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (txt2.Enabled == false) txt2.Enabled = true; else txt2.Enabled = false;
        }

        private void çıxışToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult result2 = MessageBox.Show("Davam etmək istəyirsiniz?", "?", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (result2 == DialogResult.No) { return; }

            base.Close();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            if (txtad.Text == "") { MessageBox.Show(" Soyadı A.A  Qeyd olunmayıb ... "); return; }

            DialogResult result2 = MessageBox.Show("Davam etmək istəyirsiniz?", "?", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (result2 == DialogResult.No) { return; }

                Print();
                //--------------------qaliq refres olunsun rasxod sehifede-----------------------------------------------
                cbsoyad.Text = txtad.Text;
                txtad.ForeColor = Color.DarkRed;


                MyData.selectCommand("baza.accdb", "Select * from kassaqaliq where AD = " + "'" + txtad.Text.ToString() + "'");
                MyData.dtmain = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmain);
                try
                {
                    txtqaliq.Text = Math.Round(Convert.ToDouble(MyData.dtmain.Rows[0][1]),2).ToString();
                }
                catch { };
                txtkecmisavans.Text = MyData.dtmain.Rows[0][1].ToString();

                MyData.insertCommand("baza.accdb", "insert into kassakitabi (Tarix, İstifadəçi, Mədaxil, Xərclənib, Məxaric, Qalıq, Qeyd) Values ('" + dttarix.Text + "'," + "'" + txtad.Text + "'," + "'" + "0" + "'," + "'" + "0" + "'," + "'" + Math.Round(Convert.ToDouble(txt1.Text), 2).ToString() + "'," + "'" + txtqaliq.Text + "','Kassa Qalıq - " + button19.Text + " AZN / " + richTextBox1.Text + "')");

                MyData.insertCommand("baza.accdb", "insert into Emeliyyatlar (a1, a2, a3) Values ('" + DateTime.Now + "', 'MƏXARİC - " + txtad.Text + " - " + Math.Round(Convert.ToDouble(txt1.Text),2).ToString() + " manat ( "+ richTextBox1.Text + " )','" + Environment.MachineName + "')");
                
                kassakitabi();
        }

        private void button15_Click(object sender, EventArgs e)
        {
            kassakitabi();
            QaliqRefresh();
            myrefresh2();
        }

        private void btexcel2_Click(object sender, EventArgs e)
        {
            if (cbsoyad.Text == "") { MessageBox.Show(" Soyadı A.A  Qeyd olunmayıb ... "); return; }

            DialogResult result2 = MessageBox.Show("Davam etmək istəyirsiniz?", "?", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (result2 == DialogResult.No) { return; }

                Print2();

                MyData.selectCommand("baza.accdb", "Select * from kassaqaliq where AD = " + "'" + txtad.Text.ToString() + "'");
                MyData.dtmain = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmain);
                try
                {
                    txtqaliq.Text = Math.Round(Convert.ToDouble(MyData.dtmain.Rows[0][1]),2).ToString();
                }
                catch { };

                //kassa kitabina ----------------------------------------------------------------
                MyData.insertCommand("baza.accdb", "insert into kassakitabi (Tarix, İstifadəçi, Mədaxil, Xərclənib, Məxaric, Qalıq, Qeyd) Values ('" + dttarixhesabat.Text + "'," + "'" + cbsoyad.Text + "'," + "'" + txtavansqaytarilib.Text + "'," + "'" + txtavansxerclenib.Text + "'," + "'" + "0" + "'," + "'" + txtqaliq.Text + "','Kassa Qalıq - " + button19.Text + " AZN / Verilmiş avansın bağlanması')");
                MyData.insertCommand("baza.accdb", "insert into Emeliyyatlar (a1, a2, a3) Values ('" + DateTime.Now + "', 'AVANS - " + cbsoyad.Text + " - xərclənib " + txtavansxerclenib.Text + " manat - qaytarılıb " + txtavansqaytarilib.Text + " manat','" + Environment.MachineName + "')");
                
                kassakitabi();
        }

        private void kassaSənədləriToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                File.Copy("Bos.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Kassa Senedleri.xlsx", true);
            }
            catch { MessageBox.Show("Bos.xlsx tapılmadı."); }


            int i, k;

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Kassa Senedleri.xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

            oSheet.Cells[1, 1] = "№";
            oSheet.Cells[1, 2] = "Tarix";
            oSheet.Cells[1, 3] = "İstifadəçi";
            oSheet.Cells[1, 4] = "Mədaxil";
            oSheet.Cells[1, 5] = "Xərclənib";
            oSheet.Cells[1, 6] = "Məxaric";
            oSheet.Cells[1, 7] = "Qalıq";

            for (i = 0; i < MyData.dtmain.Rows.Count; i++)
            {

                for (k = 0; k < 7; k++)
                {
                    oSheet.Cells[i+2, k+1] = MyData.dtmain.Rows[i][k].ToString();

                }

            }

            oSheet.Columns.AutoFit();
            oSheet.Rows.AutoFit();
        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {

            try
            {
                MyData.updateCommand("baza.accdb", "UPDATE kassaqaliq SET "
                                                                                     + "Qaliq ='" + (Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[6].Value) + Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[3].Value) + Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[4].Value) - Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Value)).ToString() + "'"
                                                                                     + " WHERE AD Like '" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[2].Value.ToString() + "'");

            }
            catch { MessageBox.Show("Əməliyyat baş tutmadı."); }


            try
            {
                string ST ="'" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString() + "'";
                MyData.deleteCommand("baza.accdb", "DELETE FROM kassakitabi WHERE Nomre=" + ST);

                MessageBox.Show("Tapşırıq yerinə yetirildi.");

                kassakitabi();
            }
            catch { };

        }

        private void label18_Click(object sender, EventArgs e)
        {
            if (txtqaliq.Enabled == false) { txtqaliq.Enabled = true; return; }

            txtqaliq.Enabled = false;
        }

        private void button16_Click(object sender, EventArgs e)
        {
            try
            {
                File.Copy("Rasxod.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Məxaric - " + txtad.Text + ".xlsx", true);
            }
            catch { MessageBox.Show("Rasxod.xlsx tapılmadı."); }

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Məxaric - " + txtad.Text + ".xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

            if (txt1.Text == "") txt1.Text = "0.00";
            if (txt2.Text == "") txt2.Text = "Sıfır";
            if (txtad.Text == "") txtad.Text = "Şirinov Rasim Rafiqoviç";
            oSheet.Cells[4, 7] = txtnomre.Text;
            oSheet.Cells[7, 6] = dttarix.Text;
            oSheet.Cells[28, 4] = dttarix.Text;
            oSheet.Cells[12, 4] = txtad.Text.ToString();
            oSheet.Cells[10, 6] = txt1.Text.ToString();
            oSheet.Cells[17, 3] = txt2.Text.ToString();
            oSheet.Cells[25, 3] = txt2.Text.ToString();
            oSheet.Cells[15, 4] = txtesasname.Text.ToString();

            //----------------------------------------------------------------excelde sutunlarin silinmesi 
            Excel.Range range = oSheet.get_Range("A37:AB37", Type.Missing);
            range.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

            if (txtad.Text == "Əsədov İbrahim Akif oğlu")
                oSheet.Cells[30, 3] = "Səxsiyyət Vəsiqəsi AZE № 13220118, BAKI-Nizami  RPI 27.05.2013-cü il";
            else if (txtad.Text == "Musayev Rəşad İslam oğlu")
                oSheet.Cells[30, 3] = "Səxsiyyət Vəsiqəsi AZE № 17269313 ASAN3 29.09.2017";
            else if (txtad.Text == "Rəşidov Zaur Xanəliyeviç")
                oSheet.Cells[30, 3] = "Səxsiyyət Vəsiqəsi AZE № 08272442 Səbail RPI 06.04.2011";
            else if (txtad.Text == "Heydərov Namik Abisalam oğlu")
                oSheet.Cells[30, 3] = "Səxsiyyət Vəsiqəsi AZE № 13726303 Yasamal RPI 29.03.2013";
            else if (txtad.Text == "Şirinov Rasim Rafiqoviç")
                oSheet.Cells[30, 3] = "Səxsiyyət Vəsiqəsi AZE № 08203665, BAKI-XƏTAI  RPI 16.12.2010-cu il";
            else if (txtad.Text == "Səfərli Nərmin Səfər qızı")
                oSheet.Cells[30, 3] = "Səxsiyyət Vəsiqəsi AZE № 09513223, BAKI-XƏTAI  RPI 29.08.2012-ci il";
            else if (txtad.Text == "Məmmədov İlqar İmran oğlu")
                oSheet.Cells[30, 3] = "Səxsiyyət Vəsiqəsi AZE № 09229670, SUMQAYIT ŞPİ QŞVŞ 02.09.2011-ci il";

        }

        private void button17_Click(object sender, EventArgs e)
        {
            try
            {
                File.Copy("Avans.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Avans - " + cbsoyad.Text + ".xlsx", true);
            }
            catch { MessageBox.Show("Avans.xlsx tapılmadı."); }

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Avans - " + cbsoyad.Text + ".xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

            if (txtavansalinib.Text == "") txtavansalinib.Text = "0";
            if (txtavansqaytarilib.Text == "") txtavansqaytarilib.Text = "0";
            if (txtkecmisavans.Text == "") txtkecmisavans.Text = "0";
            if (txtavansxerclenib.Text == "") txtavansxerclenib.Text = "0";

            oSheet.Cells[10, 1] = "AVANS HESABATI №  " + dttarixhesabat.Text + "-cu il.";
            if (txtavansalinib.Text != "0") oSheet.Cells[17, 2] = "( " + dttarixavansalinma.Text + " )"; else oSheet.Cells[17, 2] = "()";
            oSheet.Cells[8, 3] = cbsoyad.Text.ToString();
            oSheet.Cells[5, 6] = cbvezife.Text.ToString();
            oSheet.Cells[14, 5] = (Convert.ToDouble(txtkecmisavans.Text) - Convert.ToDouble(txtavansalinib.Text)).ToString();
            oSheet.Cells[17, 5] = txtavansalinib.Text.ToString();
            oSheet.Cells[26, 5] = txtavansxerclenib.Text.ToString();
            oSheet.Cells[27, 5] = txtavansqaytarilib.Text.ToString();

            if (textBox1.Text != "")
            {
                oSheet.Cells[41, 1] = dateTimePicker1.Text.ToString();
                oSheet.Cells[41, 3] = textBox1.Text.ToString();
                oSheet.Cells[41, 11] = txtm1.Text.ToString();
            }
            if (textBox2.Text != "")
            {
                oSheet.Cells[42, 1] = dateTimePicker2.Text.ToString();
                oSheet.Cells[42, 3] = textBox2.Text.ToString();
                oSheet.Cells[42, 11] = txtm2.Text.ToString();
            }
            if (textBox3.Text != "")
            {
                oSheet.Cells[43, 1] = dateTimePicker3.Text.ToString();
                oSheet.Cells[43, 3] = textBox3.Text.ToString();
                oSheet.Cells[43, 11] = txtm3.Text.ToString();
            }
            if (textBox4.Text != "")
            {
                oSheet.Cells[44, 1] = dateTimePicker4.Text.ToString();
                oSheet.Cells[44, 3] = textBox4.Text.ToString();
                oSheet.Cells[44, 11] = txtm4.Text.ToString();
            }
            if (textBox5.Text != "")
            {
                oSheet.Cells[45, 1] = dateTimePicker5.Text.ToString();
                oSheet.Cells[45, 3] = textBox5.Text.ToString();
                oSheet.Cells[45, 11] = txtm5.Text.ToString();
            }

        }

        private void button18_Click(object sender, EventArgs e)
        {
            DialogResult result2 = MessageBox.Show("Davam etmək istəyirsiniz?", "?", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (result2 == DialogResult.No) { return; }

            if (checkBox1.Checked == true)
            {
                MyData.updateCommand("baza.accdb", "UPDATE AGLizinqKassa  SET a1 =" + "'" + (Convert.ToDouble(button19.Text) + Convert.ToDouble(textBox8.Text + "." + textBox7.Text)).ToString() + "'");
                
                myrefresh8();

                //kassa kitabina ----------------------------------------------------------------

                MyData.insertCommand("baza.accdb", "insert into kassakitabi (Tarix, İstifadəçi, Mədaxil, Xərclənib, Məxaric, Qalıq, Qeyd) Values ('" + dttarix.Text + "','AGBANK','" + textBox8.Text + "." + textBox7.Text + "','0','0'," + "'0','Kassa Qalıq - " + button19.Text + " AZN / kassanın möhkəmləndirilməsi üçün rəsmi mədaxil')");
                

                textBox7.Text = "0";
                textBox8.Text = "0";
                kassakitabi();
                return;
            }

            MyData.updateCommand("baza.accdb", "UPDATE AGLizinqKassa  SET a1 =" + "'" + (Convert.ToDouble(button19.Text) + Convert.ToDouble(textBox8.Text + "." + textBox7.Text)).ToString() + "'" );

            textBox7.Text = "0";
            textBox8.Text = "0";
            myrefresh8();
        }

        private void əlaqəToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Elaqe elaqe = new Elaqe();
            elaqe.ShowDialog();
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView1.EditMode = DataGridViewEditMode.EditProgrammatically;

            try
            {
                MyData.updateCommand("baza.accdb", "UPDATE kassakitabi SET "
                                                                                        + "Tarix ='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString() + "',"
                                                                                     + "İstifadəçi ='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[2].Value.ToString() + "',"
                                                                                     + "Mədaxil ='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[3].Value.ToString() + "',"
                                                                                     + "Xərclənib ='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[4].Value.ToString() + "',"
                                                                                     + "Məxaric ='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Value.ToString() + "',"
                                                                                     + "Qalıq ='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[6].Value.ToString() + "',"
                                                                                     + "Qeyd ='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[7].Value.ToString() + "'"
                                                                                     + " WHERE Nomre Like '" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString() + "'");

                
                

            }
            catch { MessageBox.Show("Əməliyyat baş tutmadı.1"); }

            //kassakitabina uygun olaraq sexslerin qaliginin hesablanmasi
            double cm2=0;
            for (int cm = 0; cm < dataGridView1.Rows.Count; cm++)
            {
                if (dataGridView1.Rows[cm].Cells["İstifadəçi"].Value.ToString() == dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[2].Value.ToString())
                    cm2 += Convert.ToDouble(dataGridView1.Rows[cm].Cells[5].Value.ToString()) - Convert.ToDouble(dataGridView1.Rows[cm].Cells[4].Value.ToString()) - Convert.ToDouble(dataGridView1.Rows[cm].Cells[3].Value.ToString());
            }
                try
                {
                MyData.updateCommand("baza.accdb", "UPDATE kassaqaliq SET "
                                                                                         + "Qaliq ='" + cm2.ToString() + "'"
                                                                                         + " WHERE AD Like '" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[2].Value.ToString() + "'");

                    
                    
                    MessageBox.Show("Istifadəçinin qalığı etdiyiniz dəyişikliyə uyğun olaraq avtomatik yeniləndi.");
                    QaliqRefresh();
                }
                catch { MessageBox.Show("Əməliyyat baş tutmadı.2"); }
                
            
        }

        private void editToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.EditMode = DataGridViewEditMode.EditOnEnter;
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            try
            {
                kassakitabi();
            }
            catch { }
        }

        private void textBox6_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    kassakitabi();
                }
                catch { }
            }
        }

        private void button14_Click_2(object sender, EventArgs e)
        {
            if (richTextBox1.Enabled == true) { richTextBox1.Enabled = false; return; }
            richTextBox1.Enabled = true;
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            timer2.Enabled = true;
        }

        private void txtavansxerclenib_TextChanged(object sender, EventArgs e)
        {
            try { txtm1.Text = txtavansxerclenib.Text; if (txtavansxerclenib.Text == "0") txtm1.Text = ""; }
            catch { }
        }

        private void txtavansqaytarilib_TextChanged(object sender, EventArgs e)
        {
            try { txtm1.Text = txtavansxerclenib.Text; if (txtavansxerclenib.Text == "0") txtm1.Text = ""; }
            catch { }
        }

        private void DtBaslama_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode==Keys.Enter)
            {
                try
                {
                    kassakitabi();
                }
                catch { }
            }
        }

        private void DtBitme_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    kassakitabi();
                }
                catch { }
            }
        }
    }
}
