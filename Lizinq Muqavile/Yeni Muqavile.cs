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
using System.IO;
using System.Net;
using System.Speech.Synthesis;
using System.Globalization;
using System.Web;
using Nsoft;

namespace Lizinq_Muqavile
{
    public partial class Yeni_Muqavile : Form
    {

        Excel.Application oXL;
        Excel._Workbook oWB;
        Excel._Worksheet oSheet;

        public Yeni_Muqavile()
        {
            InitializeComponent();
        }

        public void myrefresLayiheler()
        {
            MyData.selectCommand("baza.accdb", "Select [Lizinq alan], Satıcı,[Lizinq obyekti],[Lizinq məbləği],[Lizinqin müddəti],[% dərəcəsi],Tarix from muqavilelayihe");
            MyData.dtmain = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmain);
            dataGridView1.DataSource = MyData.dtmain;
        }

        public void myrefresRekvizit()
        {
            MyData.selectCommand("baza.accdb", "Select * from muqavilerekvizit");
            MyData.dtmain = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmain);
            dataGridView2.DataSource = MyData.dtmain;
        }

        public void myrefresRekvizitSatici()
        {
            MyData.selectCommand("baza.accdb", "Select * from muqavilesaticirekvizit");
            MyData.dtmain = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmain);
            dataGridView3.DataSource = MyData.dtmain;
        }

        public void SaveRazilasma()
        {
            if (MyCheck.davamYesNo()) return;

            string s1, s2, s3, s4, s5, s6, s7, s8, s9, s10, s11, s12, s13, s14, s15, s16, s17, s18, s19, s20, s21;
           
            s1 = "'" + txtlizinqalan.Text + "'";
            s2 = "'" + txtsatici.Text + "'";
            s3 = "'" + txtobyekt.Text + "'";
            s4 = "'" + txtlizinqmebleg.Text + "'";
            s5 = "'" + txtavadanliqdeyer.Text + "'";
            s6 = "'" + txtmuddet.Text + "'";
            s7 = "'" + txtavans.Text + "'";
            s8 = "'" + txtfaiz.Text + "'";
            s9 = "'" + txtqrafik.Text + "'";
            s10 = "'" + txtguzest.Text + "'";
            s11 = "'" + txtlizmeqsedi.Text + "'";
            s12 = "'" + txtsigorta.Text + "'";
            s13 = "'" + txtteminat.Text + "'";
            s14 = "'" + txtbirdefemukafat.Text + "'";
            s15 = "'" + txtsertler.Text + "'";
            s16 = "'" + txtmonitorinq.Text + "'";
            s17 = "'" + txtnagdilasdirma.Text + "'";
            s18 = "'" + cmbkurator.Text + "'";
            s19 = "'" + dttarix.Text + "'";
            s20 = "'" + cmbfhs1.Text + "'";
            s21 = "'" + cmbfhs2.Text + "'";
           
            label4.ForeColor = Color.Black; label3.ForeColor = Color.Black; label6.ForeColor = Color.Black; label7.ForeColor = Color.Black;
            label2.ForeColor = Color.Black; label8.ForeColor = Color.Black; label9.ForeColor = Color.Black;
            label11.ForeColor = Color.Black; label12.ForeColor = Color.Black; label10.ForeColor = Color.Black; label13.ForeColor = Color.Black;
            label16.ForeColor = Color.Black; label15.ForeColor = Color.Black; label14.ForeColor = Color.Black; label17.ForeColor = Color.Black;
            label22.ForeColor = Color.Black; label23.ForeColor = Color.Black; label13.ForeColor = Color.Black; label13.ForeColor = Color.Black;
            cmbfhs1.BackColor = Color.Gainsboro; cmbfhs2.BackColor = Color.Gainsboro;

            if (txtlizinqalan.Text == "") { label4.ForeColor = Color.Red; MessageBox.Show("Lizinqalan qeyd olunmayıb ..."); return; };
            if (txtsatici.Text == "") { label3.ForeColor = Color.Red; MessageBox.Show("Satıcı qeyd olunmayıb ..."); return; };
            if (txtobyekt.Text == "") { label6.ForeColor = Color.Red; MessageBox.Show("Lizinq obyekti qeyd olunmayıb ..."); return; };
            if (txtlizinqmebleg.Text == "" || txtlizinqmebleg.Text == "0") { label7.ForeColor = Color.Red; MessageBox.Show("Lizinq məbləği qeyd olunmayıb ..."); return; };
            if (txtavadanliqdeyer.Text == "" || txtavadanliqdeyer.Text == "0") { label2.ForeColor = Color.Red; MessageBox.Show("Lizinq obyektinin dəyəri qeyd olunmayıb ..."); return; };
            if (txtmuddet.Text == "" || txtmuddet.Text == "0") { label8.ForeColor = Color.Red; MessageBox.Show("Lizinqin müddəti qeyd olunmayıb ..."); return; };
            if (txtavans.Text == "") { label9.ForeColor = Color.Red; MessageBox.Show("Avans qeyd olunmayıb ..."); return; };
            if (txtfaiz.Text == "") { label11.ForeColor = Color.Red; MessageBox.Show("Faiz dərəcəsi qeyd olunmayıb ..."); return; };
            if (txtqrafik.Text == "") { label2.ForeColor = Color.Red; MessageBox.Show("Ödəniş cədvəli qeyd olunmayıb ..."); return; };
            if (txtguzest.Text == "") { label10.ForeColor = Color.Red; MessageBox.Show("Güzəşt müddəti qeyd olunmayıb ..."); return; };
            if (txtlizmeqsedi.Text == "") { label13.ForeColor = Color.Red; MessageBox.Show("Lizinqin məqsədi qeyd olunmayıb ..."); return; };
            if (txtsigorta.Text == "") { label16.ForeColor = Color.Red; MessageBox.Show("Sigorta qeyd olunmayıb ..."); return; };
            if (txtteminat.Text == "") { label15.ForeColor = Color.Red; MessageBox.Show("Təminat qeyd olunmayıb ..."); return; };
            if (txtbirdefemukafat.Text == "") { label14.ForeColor = Color.Red; MessageBox.Show("Birdəfəlik mükafat qeyd olunmayıb ..."); return; };
            if (txtmonitorinq.Text == "") { label17.ForeColor = Color.Red; MessageBox.Show("manitorinqin vaxtdı qeyd olunmayıb ..."); return; };
            if (cmbkurator.Text == "" || cmbkurator.Text == " - ") { label22.ForeColor = Color.Red; MessageBox.Show("Təqdim edən qeyd olunmayıb ..."); return; };
            if (dttarix.Text == "") { label23.ForeColor = Color.Red; MessageBox.Show("Tarix qeyd olunmayıb..."); return; };
            if (cmbfhs1.Text == "") { cmbfhs1.BackColor = Color.Red; MessageBox.Show("Lizinq alanın növü qeyd olunmayıb..."); return; };
            if (cmbfhs2.Text == "") { cmbfhs2.BackColor = Color.Red; MessageBox.Show("Satıcının növü qeyd olunmayıb..."); return; };

                MyData.selectCommand("baza.accdb", "Select * from muqavilelayihe where [Lizinq alan] = " + "'" + txtlizinqalan.Text + "' and [Tarix] = " + "'" + dttarix.Text + "'");
                MyData.dtmain = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmain);
                try { txtkod1.Text = MyData.dtmain.Rows[0]["Код"].ToString(); }
                catch { }

                if (MyData.dtmain.Rows.Count > 0)
                {
                MyData.updateCommand("baza.accdb", "UPDATE muqavilelayihe SET "

                + "[lizinq alan] =" + s1 + "," + "Satıcı =" + s2 + "," + "[Lizinq obyekti] =" + s3 + "," + "[Lizinq məbləği] =" + s4 + ","

                + "[Avadanlığın dəyəri] =" + s5 + "," + "[Lizinqin müddəti] =" + s6 + "," + "Avans =" + s7 + "," + "[% dərəcəsi] =" + s8 + ","

                + "Qrafik =" + s9 + "," + "[Güzəşt müddəti (ay)] =" + s10 + "," + "[Lizinqin məqsədi] =" + s11 + "," + "Siğorta =" + s12 + ","

                + "Тəminat =" + s13 + "," + "[Birdəfəlik mükafat (%)] =" + s14 + "," + "[Cari şərtlər] =" + s15 + "," + "[İlkin monitorinq] =" + s16 + ","

                + "[Nağdılaşdırma (%)] =" + s17 + "," + "Kurator =" + s18 + "," + "Tarix =" + s19 + "," + "[Müştəri növü] =" + s20 + "," + "[Satıcı növü] =" + s21

                + "WHERE Код Like" + "'" + txtkod1.Text + "'");

                    MessageBox.Show("Dəyişiklik olundu..");
                }
                else
                {
                MyData.insertCommand("baza.accdb", "insert into muqavilelayihe ([lizinq alan],Satıcı,[Lizinq obyekti],[Lizinq məbləği],[Avadanlığın dəyəri],[Lizinqin müddəti],Avans,[% dərəcəsi],Qrafik,[Güzəşt müddəti (ay)],[Lizinqin məqsədi],Siğorta,Тəminat,[Birdəfəlik mükafat (%)],[Cari şərtlər],[İlkin monitorinq],[Nağdılaşdırma (%)],Kurator,Tarix,[Müştəri növü],[Satıcı növü])values("

                + s1 + "," + s2 + "," + s3 + "," + s4 + "," + s5 + "," + s6 + "," + s7 + "," + s8 + "," + s9 + "," + s10 + "," + s11 + ","
                + s12 + "," + s13 + "," + s14 + "," + s15 + "," + s16 + "," + s17 + "," + s18 + "," + s19 + "," + s20 + "," + s21 + ")");

                    MessageBox.Show("Yeni məlumat bazaya əlavə edildi");
                }
        }

        public void SaveLizinqalanRekvizit()
        {
            if (MyCheck.davamYesNo()) return;
            string s1, s2, s3, s4, s5, s6, s7, s8, s9, s10, s11, s12, s13, s14, s15;

            s1 = "'" + textBox34.Text + "'"; //----Satici
            s2 = "'" + txtfizsexsiyyetseriya.Text + " № " + txtfizsexsiyyetnomre.Text + "'"; // ------------------sexsiyyet seriya nomre
            s3 = "'" + txtfizqeydiyyat.Text + "'"; //------------------qeydiyyat unvani
            s4 = "'" + txtfizfaktikiunvan.Text + "'"; //--------------------faktiki unvan
            s5 = "'" + txtfizvesiqeveren.Text + "'"; //-------------vesiqe veren orqan
            s6 = "'" + dtfizvesverilmetarix.Text + "'"; //-----------------verilme tarixi
            s7 = "'" + cmbfiztelkod1.Text + " " + txtfiznomre1.Text + "'"; //----nomre1
            s8 = "'" + cmbfiztelkod2.Text + " " + txtfiznomre2.Text + "'"; //----nomre2
            s9 = "'" + cmbfiztelkod3.Text + " " + txtfiznomre3.Text + "'"; //----nomre3
            s10 = "'" + txtfizrekvizit.Text + "'"; //----rekvizitler
            s11 = "'" + txtVoen1.Text + "'"; //-----------------Voen Nomre
            s12 = "'" + txtVoenHuquqiUnvan1.Text + "'"; //----HuqUnvan
            s13 = "'" + txtVoenVerenOrqan1.Text + "'"; //----Unvan
            s14 = "'" + txtVoenTarix1.Text + "'"; //----Tarix
            s15 = "'" + txtDirektor1.Text + "'"; //----Direktor

            if (textBox34.Text == "") { MessageBox.Show("Satıcının adı qeyd olunmayıb !!!"); return; }

            MyData.selectCommand("baza.accdb", "Select * from muqavilerekvizit where [Şəxsiyyət vəsiqəsi] Like " + "'" + txtfizsexsiyyetseriya.Text + " № " + txtfizsexsiyyetnomre.Text + "'");
            MyData.dtmain = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmain);

            if (MyData.dtmain.Rows.Count > 0)
            {
                MyData.updateCommand("baza.accdb", "UPDATE muqavilerekvizit SET "

                + "[Lizinq alan] =" + s1 + "," + "[Şəxsiyyət vəsiqəsi] =" + s2 + "," + "[Ş/V qeydiyyat ünvanı] =" + s3 + ","

                + "[Ş/V faktiki ünvan] =" + s4 + "," + "[Ş/V verən orqan] =" + s5 + "," + "[Ş/V verilmə tarixi] =" + s6 + ","

                + "[Əlaqə nömrəsi 1] =" + s7 + "," + "[Əlaqə nömrəsi 2] =" + s8 + "," + "[Əlaqə nömrəsi 3] =" + s9 + "," + "[Rekvizitər] =" + s10

                + " WHERE [Şəxsiyyət vəsiqəsi] like" + s2);

                MessageBox.Show("Dəyişiklik olundu..");
            }
            else
            {
                if (voenLizinqalan.Visible == false)
                {
                    MyData.insertCommand("baza.accdb", "insert into muqavilerekvizit ([Lizinq alan],[Şəxsiyyət vəsiqəsi],[Ş/V qeydiyyat ünvanı],[Ş/V faktiki ünvan],[Ş/V verən orqan],[Ş/V verilmə tarixi],[Əlaqə nömrəsi 1],[Əlaqə nömrəsi 2],[Əlaqə nömrəsi 3],Rekvizitər)values("

                    + s1 + "," + s2 + "," + s3 + "," + s4 + "," + s5 + "," + s6 + "," + s7 + "," + s8 + "," + s9 + "," + s10 + ")");

                    MessageBox.Show("Yeni məlumat bazaya əlavə edildi");
                }
                else
                {
                    MyData.insertCommand("baza.accdb", "insert into muqavilerekvizit ([Lizinq alan],[Şəxsiyyət vəsiqəsi],[Ş/V qeydiyyat ünvanı],[Ş/V faktiki ünvan],[Ş/V verən orqan],[Ş/V verilmə tarixi],[Əlaqə nömrəsi 1],[Əlaqə nömrəsi 2],[Əlaqə nömrəsi 3],Rekvizitər,Vöen,[Vöen qeydiyyat ünvanı],[Vöen verən orqan],[Vöen verilmə tarixi],Direktor)values("

                    + s1 + "," + s2 + "," + s3 + "," + s4 + "," + s5 + "," + s6 + "," + s7 + "," + s8 + "," + s9 + "," + s10 + "," + s11 + "," + s12 + "," + s13 + "," + s14 + "," + s15 + ")");

                    MessageBox.Show("Yeni məlumat bazaya əlavə edildi");
                }
            } 
        }

        public void SaveSaticiRekvizit()
        {
            if (MyCheck.davamYesNo()) return;

            string s1, s2, s3, s4, s5, s6, s7, s8, s9, s10, s11, s12, s13, s14, s15;

            s1 = "'" + txtsatici2.Text + "'"; //----Satici
            s2 = "'" + textBox21.Text + " № " + textBox20.Text + "'"; // ------------------sexsiyyet seriya nomre
            s3 = "'" + textBox17.Text + "'"; //------------------qeydiyyat unvani
            s4 = "'" + textBox18.Text + "'"; //--------------------faktiki unvan
            s5 = "'" + textBox19.Text + "'"; //-------------vesiqe veren orqan
            s6 = "'" + dateTimePicker2.Text + "'"; //-----------------verilme tarixi
            s7 = "'" + comboBox3.Text + " " + maskedTextBox3.Text + "'"; //----nomre1
            s8 = "'" + comboBox2.Text + " " + maskedTextBox3.Text + "'"; //----nomre2
            s9 = "'" + comboBox1.Text + " " + maskedTextBox3.Text + "'"; //----nomre3
            s10 = "'" + richTextBox1.Text + "'"; //----rekvizitler
            s11 = "'" + txtVoen2.Text + "'"; //-----------------Voen Nomre
            s12 = "'" + txtVoenHuquqiUnvan2.Text + "'"; //----HuqUnvan
            s13 = "'" + txtVoenVerenOrqan2.Text + "'"; //----Unvan
            s14 = "'" + txtVoenTarix2.Text + "'"; //----Tarix
            s15 = "'" + txtDirektor2.Text + "'"; //----Direktor
            
            if (txtsatici2.Text == "") { MessageBox.Show("Satıcının adı qeyd olunmayıb !!!"); return; }

            
            
            
            MyData.selectCommand("baza.accdb", "Select * from muqavilesaticirekvizit where [Şəxsiyyət vəsiqəsi] Like " + "'" + textBox21.Text + " № " + textBox20.Text + "'");
            MyData.dtmain = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmain);

            if (MyData.dtmain.Rows.Count > 0)
            {
                MyData.updateCommand("baza.accdb", "UPDATE muqavilesaticirekvizit SET "
                + "[Satıcı] =" + s1 + "," + "[Şəxsiyyət vəsiqəsi] =" + s2 + "," + "[Ş/V qeydiyyat ünvanı] =" + s3 + ","
                + "[Ş/V faktiki ünvan] =" + s4 + "," + "[Ş/V verən orqan] =" + s5 + "," + "[Ş/V verilmə tarixi] =" + s6 + ","
                + "[Əlaqə nömrəsi 1] =" + s7 + "," + "[Əlaqə nömrəsi 2] =" + s8 + "," + "[Əlaqə nömrəsi 3] =" + s9 + "," + "[Rekvizitər] =" + s10
                + " WHERE [Şəxsiyyət vəsiqəsi] =" + s2);

                MessageBox.Show("Dəyişiklik olundu..");
            }
            else
            {
                if (VoenSatici.Visible == false)
                {



                    MyData.insertCommand("baza.accdb", "insert into muqavilesaticirekvizit (Satıcı,[Şəxsiyyət vəsiqəsi],[Ş/V qeydiyyat ünvanı],[Ş/V faktiki ünvan],[Ş/V verən orqan],[Ş/V verilmə tarixi],[Əlaqə nömrəsi 1],[Əlaqə nömrəsi 2],[Əlaqə nömrəsi 3],Rekvizitər)values("

                    + s1 + "," + s2 + "," + s3 + "," + s4 + "," + s5 + "," + s6 + "," + s7 + "," + s8 + "," + s9 + "," + s10 + ")");
                    MessageBox.Show("Yeni məlumat bazaya əlavə edildi");
                }
                else
                {
                    MyData.insertCommand("baza.accdb", "insert into muqavilesaticirekvizit (Satıcı,[Şəxsiyyət vəsiqəsi],[Ş/V qeydiyyat ünvanı],[Ş/V faktiki ünvan],[Ş/V verən orqan],[Ş/V verilmə tarixi],[Əlaqə nömrəsi 1],[Əlaqə nömrəsi 2],[Əlaqə nömrəsi 3],Rekvizitər,Vöen,[Vöen qeydiyyat ünvanı],[Vöen verən orqan],[Vöen verilmə tarixi],Direktor)values("

                    + s1 + "," + s2 + "," + s3 + "," + s4 + "," + s5 + "," + s6 + "," + s7 + "," + s8 + "," + s9 + "," + s10 + "," + s11 + "," + s12 + "," + s13 + "," + s14 + "," + s15 + ")");
                    MessageBox.Show("Yeni məlumat bazaya əlavə edildi");
                }
            } 
        }

        public void SaveNeqliyyat()
        {
            if (MyCheck.davamYesNo()) return;

            string s1, s2, s3, s4, s5, s6, s7, s8, s9, s10, s11, s12, s13;
            
            s1 = "'" + txtnomre1.Text + "'";
            s2 = "'" + txtbannomre2.Text + "'";
            s3 = "'" + txtmuherriknomre3.Text + "'";
            s4 = "'" + txtsassinomre4.Text + "'";
            s5 = "'" + txttexpass5.Text + "'";
            s6 = "'" + txtburaxilisili6.Text + "'";
            s7 = "'" + txtpassverilmetarix7.Text + "'";
            s8 = "'" + txtmarkasi8.Text + "'";
            s9 = "'" + txtlizinqalanadi9.Text + "'";
            s10 = "'" + txtlayihe.Text + "'";
            s11 = "'" + txtzavodnomresi11.Text + "'";
            s12 = "'" + txtreng12.Text + "'";
            s13 = "'" + txtQeyd.Text + "'";

            MyData.selectCommand("baza.accdb", "Select * from etibarnameneqliyyat where c8=" + s2 + "and c9=" + s3 + "and c10=" + s4); ;
            MyData.dtmain = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmain);

            if (MyData.dtmain.Rows.Count > 0)
            {
                MyData.updateCommand("baza.accdb", "UPDATE etibarnameneqliyyat SET "
                + "c1 =" + s1 + "," + "c8 =" + s2 + "," + "c9 =" + s3 + "," + "c10 =" + s4 + "," + "c12 =" + s5 + "," + "c6 =" + s6 + ","
                + "c7 =" + s7 + "," + "c2 =" + s8 + "," + "c3 =" + s9 + "," + "c4 =" + s10 + "," + "c11 =" + s11 + "," + "c5 =" + s12 + "," + "c13 =" + s13
                + " WHERE c8=" + s2 + "and c9=" + s3 + "and c10=" + s4);

                MessageBox.Show("Dəyişiklik olundu...");

            }
            else
            {
                MyData.insertCommand("baza.accdb", "insert into etibarnameneqliyyat (c1,c8,c9,c10,c12,c6,c7,c2,c3,c4,c11,c5,c13)values("
                  + "'" + txtnomre1.Text + "','" + txtbannomre2.Text + "','" + txtmuherriknomre3.Text + "','" + txtsassinomre4.Text + "','"
                  + txttexpass5.Text + "','" + txtburaxilisili6.Text + "','" + txtpassverilmetarix7.Text + "','" + txtmarkasi8.Text + "','"
                  + txtlizinqalanadi9.Text + "','" + txtlayihe.Text + "','" + txtzavodnomresi11.Text + "','" + txtreng12.Text + "','" + txtQeyd.Text + "')");

                MessageBox.Show("Yeni məlumat əlavə edildi");
            }
        }
     
        public void PrintDaxiliMaliyye()
        {
            try { File.Copy("Daxili maliyye.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Daxili maliyye - " + txtlizinqalan.Text + ".doc", true); }
            catch { MessageBox.Show("'Daxili maliyye.doc' tapılmadı."); }

            object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\Daxili maliyye - " + txtlizinqalan.Text + ".doc";

            string umumimebleg, mukafat, umumiXercler;
            double k, k2, k3, k4 = 1, k5, k6;
            int s;
            k = Convert.ToDouble(txtfaiz.Text) / 100 / 12;
            k2 = Convert.ToDouble(txtmuddet.Text);
            k3 = Convert.ToDouble(txtlizinqmebleg.Text);
            k5 = Convert.ToDouble(txtbirdefemukafat.Text) / 100 * Convert.ToDouble(txtlizinqmebleg.Text); ;
            k6 = Convert.ToDouble(txtsigorta.Text) + Convert.ToDouble(txtavans.Text) + Convert.ToDouble(txtavadanliqdeyer.Text) * Convert.ToDouble(txtnagdilasdirma.Text)/100;
            for (s = 0; s < k2; s++) { k4 = k4 * (1 + k); }
            umumimebleg = Math.Round((k * k4 / (k4 - 1) * k3 * k2 + k5 + k6), 2).ToString(); txtumumimeblegReqem.Text = umumimebleg;
            mukafat = Math.Round((k * k4 / (k4 - 1) * k3 * k2 - k3 + k5), 2).ToString(); txtmukafatReqem.Text = mukafat;
            umumiXercler = Math.Round(Convert.ToDouble((Convert.ToDouble(txtavans.Text) + Convert.ToDouble(txtnagdilasdirma2.Text) + Convert.ToDouble(txtsigorta.Text) + Convert.ToDouble(txtbirdefemukafat2.Text)).ToString()), 2).ToString(); txtXerclerReqem.Text = umumiXercler;
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


            string b = "";

            if (dttarix.Text.Substring(3, 2) == "01") b = "Yanvar";
            if (dttarix.Text.Substring(3, 2) == "02") b = "Fevral";
            if (dttarix.Text.Substring(3, 2) == "03") b = "Mart";
            if (dttarix.Text.Substring(3, 2) == "04") b = "Aprel";
            if (dttarix.Text.Substring(3, 2) == "05") b = "May";
            if (dttarix.Text.Substring(3, 2) == "06") b = "İyun";
            if (dttarix.Text.Substring(3, 2) == "07") b = "İyul";
            if (dttarix.Text.Substring(3, 2) == "08") b = "Avqust";
            if (dttarix.Text.Substring(3, 2) == "09") b = "Sentyabr";
            if (dttarix.Text.Substring(3, 2) == "10") b = "Oktyabr";
            if (dttarix.Text.Substring(3, 2) == "11") b = "Noyabr";
            if (dttarix.Text.Substring(3, 2) == "12") b = "Dekabr";

            MyChange.FindAndReplace(word, "000", txtlayihe.Text);
            MyChange.FindAndReplace(word, "111", dttarix.Text.Substring(0,2) + " " + b + " " + dttarix.Text.Substring(dttarix.Text.Length-4, 4) + "- cı il");
            MyChange.FindAndReplace(word, "222", txtlizinqalan.Text);
            MyChange.FindAndReplace(word, "222", txtlizinqalan.Text);
            MyChange.FindAndReplace(word, "222", txtlizinqalan.Text);
            MyChange.FindAndReplace(word, "333", txtsatici.Text);
            MyChange.FindAndReplace(word, "444", umumimebleg + " (" + txtumumimeblegHerf.Text + ") " );
            MyChange.FindAndReplace(word, "555", mukafat + " (" + txtmukafatHerf.Text + ") ");
            MyChange.FindAndReplace(word, "666", txtbirdefemukafat.Text);
            MyChange.FindAndReplace(word, "777", txtmuddet.Text + " (" + txtherf4.Text + ")");
            MyChange.FindAndReplace(word, "888", txtavans.Text + " (" + txtherf3.Text + ")");
            MyChange.FindAndReplace(word, "999", Math.Round(Convert.ToDouble(txtnagdilasdirma2.Text), 2).ToString() + " (" + txtherf8.Text + ")");
            MyChange.FindAndReplace(word, "1111", txtsigorta.Text + " (" + txtherf7.Text + ")");
            MyChange.FindAndReplace(word, "2222", Math.Round(Convert.ToDouble(txtbirdefemukafat2.Text), 2).ToString() + " (" + txtherf6.Text + ")");
            MyChange.FindAndReplace(word, "3333", umumiXercler + " (" + txtXerclerHerf.Text + ")");
            MyChange.FindAndReplace(word, "4444", txtDebbeFaiziReqem.Text + " (" + txtDebbeFaiziHerf.Text + ") ");
            MyChange.FindAndReplace(word, "5555", txtfizrekvizit.Text);
            MyChange.FindAndReplace(word, "6666", "Ş/V " + txtfizsexsiyyetseriya.Text + " № " + txtfizsexsiyyetnomre.Text);

            doc.Save();
        }

        public void PrintElave1()
        {
            try { File.Copy("Daxili maliyye əlavə1.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\əlavə1 - " + txtlizinqalan.Text + ".doc", true); }
            catch { MessageBox.Show("'Daxili maliyye əlavə1.doc' tapılmadı."); }

            object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\əlavə1 - " + txtlizinqalan.Text + ".doc";

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


            string b = "";

            if (dttarix.Text.Substring(3, 2) == "01") b = "Yanvar";
            if (dttarix.Text.Substring(3, 2) == "02") b = "Fevral";
            if (dttarix.Text.Substring(3, 2) == "03") b = "Mart";
            if (dttarix.Text.Substring(3, 2) == "04") b = "Aprel";
            if (dttarix.Text.Substring(3, 2) == "05") b = "May";
            if (dttarix.Text.Substring(3, 2) == "06") b = "İyun";
            if (dttarix.Text.Substring(3, 2) == "07") b = "İyul";
            if (dttarix.Text.Substring(3, 2) == "08") b = "Avqust";
            if (dttarix.Text.Substring(3, 2) == "09") b = "Sentyabr";
            if (dttarix.Text.Substring(3, 2) == "10") b = "Oktyabr";
            if (dttarix.Text.Substring(3, 2) == "11") b = "Noyabr";
            if (dttarix.Text.Substring(3, 2) == "12") b = "Dekabr";

            MyChange.FindAndReplace(word, "000", txtlayihe.Text);
            MyChange.FindAndReplace(word, "000", txtlayihe.Text);
            MyChange.FindAndReplace(word, "000", txtlayihe.Text);
            MyChange.FindAndReplace(word, "000", txtlayihe.Text);
            MyChange.FindAndReplace(word, "111", dttarix.Text.Substring(0, 2) + " " + b + " " + dttarix.Text.Substring(dttarix.Text.Length - 4, 4) + "- cı il");
            MyChange.FindAndReplace(word, "111", dttarix.Text.Substring(0, 2) + " " + b + " " + dttarix.Text.Substring(dttarix.Text.Length - 4, 4) + "- cı il");
            MyChange.FindAndReplace(word, "111", dttarix.Text.Substring(0, 2) + " " + b + " " + dttarix.Text.Substring(dttarix.Text.Length - 4, 4) + "- cı il");
            MyChange.FindAndReplace(word, "111", dttarix.Text.Substring(0, 2) + " " + b + " " + dttarix.Text.Substring(dttarix.Text.Length - 4, 4) + "- cı il");
            MyChange.FindAndReplace(word, "111", dttarix.Text.Substring(0, 2) + " " + b + " " + dttarix.Text.Substring(dttarix.Text.Length - 4, 4) + "- cı il");
            MyChange.FindAndReplace(word, "222", txtlizinqalan.Text);
            MyChange.FindAndReplace(word, "222", txtlizinqalan.Text);
            MyChange.FindAndReplace(word, "333", txtobyekt.Text);
            MyChange.FindAndReplace(word, "444", txtburaxilisili6.Text);
            MyChange.FindAndReplace(word, "555", "Ş/V " + txtfizsexsiyyetseriya.Text + " № " + txtfizsexsiyyetnomre.Text);
            MyChange.FindAndReplace(word, "666", txtbannomre2.Text);
            MyChange.FindAndReplace(word, "777", txtmuherriknomre3.Text);
            MyChange.FindAndReplace(word, "888", txtsassinomre4.Text);
            MyChange.FindAndReplace(word, "999", txtreng12.Text);
          
            doc.Save();
        }

        public void PrintYekunSifaris()
        {
            try { File.Copy("Yekun sifariş.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Yekun sifariş - " + txtlizinqalan.Text + ".doc", true); }
            catch { MessageBox.Show("'Yekun sifariş.doc' tapılmadı."); }

            object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\Yekun sifariş - " + txtlizinqalan.Text + ".doc";

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


            string b = "";

            if (dttarix.Text.Substring(3, 2) == "01") b = "Yanvar";
            if (dttarix.Text.Substring(3, 2) == "02") b = "Fevral";
            if (dttarix.Text.Substring(3, 2) == "03") b = "Mart";
            if (dttarix.Text.Substring(3, 2) == "04") b = "Aprel";
            if (dttarix.Text.Substring(3, 2) == "05") b = "May";
            if (dttarix.Text.Substring(3, 2) == "06") b = "İyun";
            if (dttarix.Text.Substring(3, 2) == "07") b = "İyul";
            if (dttarix.Text.Substring(3, 2) == "08") b = "Avqust";
            if (dttarix.Text.Substring(3, 2) == "09") b = "Sentyabr";
            if (dttarix.Text.Substring(3, 2) == "10") b = "Oktyabr";
            if (dttarix.Text.Substring(3, 2) == "11") b = "Noyabr";
            if (dttarix.Text.Substring(3, 2) == "12") b = "Dekabr";

            MyChange.FindAndReplace(word, "000", txtlayihe.Text);
            MyChange.FindAndReplace(word, "111", dttarix.Text.Substring(0, 2) + " " + b + " " + dttarix.Text.Substring(dttarix.Text.Length - 4, 4) + "- cı il");
            MyChange.FindAndReplace(word, "222", txtlizinqalan.Text);
            MyChange.FindAndReplace(word, "222", txtlizinqalan.Text + " (Ş/V " + txtfizsexsiyyetseriya.Text + " № " + txtfizsexsiyyetnomre.Text + ")");
            MyChange.FindAndReplace(word, "222", txtlizinqalan.Text);
            MyChange.FindAndReplace(word, "222", txtlizinqalan.Text);
            MyChange.FindAndReplace(word, "333", txtobyekt.Text);
            MyChange.FindAndReplace(word, "444", txtavadanliqdeyer.Text + " (" + txtherf2.Text + ")");
            MyChange.FindAndReplace(word, "555", txtsatici.Text);
            MyChange.FindAndReplace(word, "666", txtavans.Text + " (" + txtherf3.Text + ")");
            MyChange.FindAndReplace(word, "777", txtfizqeydiyyat.Text);
            MyChange.FindAndReplace(word, "888", cmbfiztelkod1.Text + " " + txtfiznomre1.Text);
            MyChange.FindAndReplace(word, "888", cmbfiztelkod1.Text + " " + txtfiznomre1.Text);
            MyChange.FindAndReplace(word, "888", cmbfiztelkod1.Text + " " + txtfiznomre1.Text);
            MyChange.FindAndReplace(word, "999", cmbfiztelkod2.Text + " " + txtfiznomre2.Text);
            MyChange.FindAndReplace(word, "999", cmbfiztelkod2.Text + " " + txtfiznomre2.Text);
            MyChange.FindAndReplace(word, "999", cmbfiztelkod2.Text + " " + txtfiznomre2.Text);
            MyChange.FindAndReplace(word, "1111", cmbfiztelkod3.Text + " " + txtfiznomre3.Text);

            doc.Save();
        }

        public void PrintQrafik()      //--------------------Print qrafik---------------------------------------
        {
            try { File.Copy("Qrafik.xlsm", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Qrafik - " + txtlizinqalan.Text + ".xlsm", true); }
            catch { MessageBox.Show("'Yekun sifariş.doc' tapılmadı."); }

            DateTime dt = dttarix.Value.Date;

            int s = 0, s3 = 0, s2s = 0, s3s = 0, s4s = 0, s5s = 0;

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Qrafik - " + txtlizinqalan.Text + ".xlsm"));
            oSheet = (Excel._Worksheet)oWB.Sheets[2];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

            oSheet.Cells[5, 3] = txtlayihe.Text;
            oSheet.Cells[6, 3] = dt.ToString("dd.MM.yy");

            if (txtmuddet.Text == "12") oSheet.Cells[13, 3] = dt.AddMonths(12).ToString("dd.MM.yy");
            else if (txtmuddet.Text == "24") oSheet.Cells[13, 3] = dt.AddMonths(24).ToString("dd.MM.yy");
            else if (txtmuddet.Text == "36") oSheet.Cells[13, 3] = dt.AddMonths(36).ToString("dd.MM.yy");
            else if (txtmuddet.Text == "48") oSheet.Cells[13, 3] = dt.AddMonths(48).ToString("dd.MM.yy");
            else if (txtmuddet.Text == "60") oSheet.Cells[13, 3] = dt.AddMonths(60).ToString("dd.MM.yy");

            oSheet.Cells[7, 3] = txtmuddet.Text;
            oSheet.Cells[8, 3] = (Convert.ToDouble(txtfaiz.Text) / 100).ToString();
            oSheet.Cells[10, 3] = txtguzest.Text;
            oSheet.Cells[12, 3] = dt.AddMonths(1).ToString("dd.MM.yy");

            oSheet.Cells[15, 3] = (100 - Convert.ToDouble(txtlizinqmebleg.Text) * 100 / Convert.ToDouble(txtavadanliqdeyer.Text)) / 100;
            oSheet.Cells[16, 3] = Convert.ToDouble(txtbirdefemukafat.Text) / 100;
            oSheet.Cells[21, 3] = txtavadanliqdeyer.Text;
            oSheet.Cells[30, 3] = Convert.ToDouble(txtavadanliqdeyer.Text) * Convert.ToDouble(txtnagdilasdirma.Text) / 100;
            oSheet.Cells[38, 3] = Convert.ToDouble(txtsigorta.Text);

            for (s = 0; s < txtlizinqalan.Text.Length; s++)
            {
                if (txtlizinqalan.Text.Substring(s, 1) == " ") s3 = s3 + 1;
                if (s3 == 0) s2s = s;
                if (s3 == 1) s3s = s;
                if (s3 == 2) s4s = s;
                if (s3 == 3) s5s = s;
            }
            try
            {
                oSheet.Cells[51, 3] = txtlizinqalan.Text.Substring(s2s + 2, s3s - s2s - 1);
                oSheet.Cells[52, 3] = txtlizinqalan.Text.Substring(0, s2s + 1);
                oSheet.Cells[53, 3] = txtlizinqalan.Text.Substring(s3s + 2, s4s - s3s - 1);
            }
            catch { };

            //**************************monitorinqin tarixinin yazilması*************************************************
            oSheet = (Excel._Worksheet)oWB.Sheets[3];
            oSheet.Activate();
            oSheet.Range["A1"].Select();

            oXL.Visible = true;
            oXL.DisplayAlerts = false; 
            oWB.Save();
        }

        public void PrintAlqiSatqi()
        {
            try { File.Copy("Alqi Satqi.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Alqi Satqi - " + txtlizinqalan.Text + ".doc", true); }
            catch { MessageBox.Show("'Alqi Satqi.doc' tapılmadı."); }

            object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\Alqi Satqi - " + txtlizinqalan.Text + ".doc";

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

            MyChange.FindAndReplace(word, "000", txtlayihe.Text);
            MyChange.FindAndReplace(word, "000", txtlayihe.Text);
            MyChange.FindAndReplace(word, "000", txtlayihe.Text);
            MyChange.FindAndReplace(word, "000", txtlayihe.Text);
            MyChange.FindAndReplace(word, "000", txtlayihe.Text);
            MyChange.FindAndReplace(word, "000", txtlayihe.Text);
            MyChange.FindAndReplace(word, "000", txtlayihe.Text);
            MyChange.FindAndReplace(word, "111", dt.Day + " " + b + " " + dt.Year + "- ci il");
            MyChange.FindAndReplace(word, "111", dt.Day  + " " + b + " " + dt.Year + "- ci il");
            MyChange.FindAndReplace(word, "222", txtlizinqalan.Text);
            MyChange.FindAndReplace(word, "222", txtlizinqalan.Text);
            MyChange.FindAndReplace(word, "222", txtlizinqalan.Text);
            MyChange.FindAndReplace(word, "333", txtsatici.Text);
            MyChange.FindAndReplace(word, "333", txtsatici.Text);
            MyChange.FindAndReplace(word, "444", txtobyekt.Text);
            MyChange.FindAndReplace(word, "555", txtavadanliqdeyer.Text + " (" + txtherf2.Text + ") ");
            MyChange.FindAndReplace(word, "555", txtavadanliqdeyer.Text + " (" + txtherf2.Text + ") ");
            MyChange.FindAndReplace(word, "666", richTextBox1.Text);
            MyChange.FindAndReplace(word, "777", txtfizrekvizit.Text);
            MyChange.FindAndReplace(word, "888", "Ş/V " + textBox21.Text + " № " + textBox20.Text);
            MyChange.FindAndReplace(word, "999", "Ş/V " + txtfizsexsiyyetseriya.Text + " № " + txtfizsexsiyyetnomre.Text);
           
            doc.Save();
        }

        public void PrintAlqiSatqiElave1()
        {
            try { File.Copy("Alqi Satqi Elave1.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Alqi Satqi Elave1 - " + txtlizinqalan.Text + ".doc", true); }
            catch { MessageBox.Show("'Alqi Satqi Elave1.doc' tapılmadı."); }

            object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\Alqi Satqi Elave1 - " + txtlizinqalan.Text + ".doc";

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

            MyChange.FindAndReplace(word, "000", txtlayihe.Text.Substring(0, 1) + "P" + txtlayihe.Text.Substring(1, txtlayihe.Text.Length - 1));
            MyChange.FindAndReplace(word, "000", txtlayihe.Text.Substring(0, 1) + "P" + txtlayihe.Text.Substring(1, txtlayihe.Text.Length - 1));
            MyChange.FindAndReplace(word, "000", txtlayihe.Text.Substring(0, 1) + "P" + txtlayihe.Text.Substring(1, txtlayihe.Text.Length - 1));
            MyChange.FindAndReplace(word, "000", txtlayihe.Text.Substring(0, 1) + "P" + txtlayihe.Text.Substring(1, txtlayihe.Text.Length - 1));
            MyChange.FindAndReplace(word, "000", txtlayihe.Text.Substring(0, 1) + "P" + txtlayihe.Text.Substring(1, txtlayihe.Text.Length - 1));
            MyChange.FindAndReplace(word, "111", dt.Day + " " + b + " " + dt.Year + "- ci il");
            MyChange.FindAndReplace(word, "111", dt.Day + " " + b + " " + dt.Year + "- ci il");
            MyChange.FindAndReplace(word, "111", dt.Day + " " + b + " " + dt.Year + "- ci il");
            MyChange.FindAndReplace(word, "111", dt.Day + " " + b + " " + dt.Year + "- ci il");
            MyChange.FindAndReplace(word, "111", dt.Day + " " + b + " " + dt.Year + "- ci il");
            MyChange.FindAndReplace(word, "222", txtlizinqalan.Text);
            MyChange.FindAndReplace(word, "222", txtlizinqalan.Text);
            MyChange.FindAndReplace(word, "333", txtsatici.Text);
            MyChange.FindAndReplace(word, "333", txtsatici.Text);
            MyChange.FindAndReplace(word, "444", txtmarkasi8.Text);
            MyChange.FindAndReplace(word, "555", txtburaxilisili6.Text);
            MyChange.FindAndReplace(word, "666", txtavadanliqdeyer.Text+ " (" + txtherf2.Text + ") ");
            MyChange.FindAndReplace(word, "666", txtavadanliqdeyer.Text + " (" + txtherf2.Text + ") ");
            MyChange.FindAndReplace(word, "666", txtavadanliqdeyer.Text + " (" + txtherf2.Text + ") ");
            MyChange.FindAndReplace(word, "777", "Ş/V " + textBox21.Text + " № " + textBox20.Text);
            MyChange.FindAndReplace(word, "888", "Ş/V " + txtfizsexsiyyetseriya.Text + " № " + txtfizsexsiyyetnomre.Text);

            doc.Save();
        }

        public void PrintQebulİstifade()
        {
            try { File.Copy("Qebul Istifade.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Qəbul İstifadə - " + txtlizinqalan.Text + ".doc", true); }
            catch { MessageBox.Show("'Qəbul İstifadə.doc' tapılmadı."); }

            object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\Qəbul İstifadə - " + txtlizinqalan.Text + ".doc";

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

            MyChange.FindAndReplace(word, "000", txtlayihe.Text);
            MyChange.FindAndReplace(word, "000", txtlayihe.Text);
            MyChange.FindAndReplace(word, "000", txtlayihe.Text);
            MyChange.FindAndReplace(word, "000", txtlayihe.Text);
            MyChange.FindAndReplace(word, "111", dt.Day + " " + b + " " + dt.Year + "- ci il");
            MyChange.FindAndReplace(word, "111", dt.Day + " " + b + " " + dt.Year + "- ci il");
            MyChange.FindAndReplace(word, "111", dt.Day + " " + b + " " + dt.Year + "- ci il");
            MyChange.FindAndReplace(word, "111", dt.Day + " " + b + " " + dt.Year + "- ci il");
            MyChange.FindAndReplace(word, "111", dt.Day + " " + b + " " + dt.Year + "- ci il");
            MyChange.FindAndReplace(word, "222", txtlizinqalan.Text);
            MyChange.FindAndReplace(word, "222", txtlizinqalan.Text);
            MyChange.FindAndReplace(word, "222", txtlizinqalan.Text);
            MyChange.FindAndReplace(word, "333", txtsatici.Text);
            MyChange.FindAndReplace(word, "333", txtsatici.Text);
            MyChange.FindAndReplace(word, "444", txtmarkasi8.Text);
            MyChange.FindAndReplace(word, "555", txtburaxilisili6.Text);
            MyChange.FindAndReplace(word, "666", txtavadanliqdeyer.Text + " (" + txtherf2.Text + ") ");
            MyChange.FindAndReplace(word, "666", txtavadanliqdeyer.Text + " (" + txtherf2.Text + ") ");
            MyChange.FindAndReplace(word, "666", txtavadanliqdeyer.Text + " (" + txtherf2.Text + ") ");
            MyChange.FindAndReplace(word, "777", "Ş/V " + textBox21.Text + " № " + textBox20.Text);
            MyChange.FindAndReplace(word, "888", "Ş/V " + txtfizsexsiyyetseriya.Text + " № " + txtfizsexsiyyetnomre.Text);
            MyChange.FindAndReplace(word, "999", txtnomre1.Text);
            MyChange.FindAndReplace(word, "1111", txttexpass5.Text);
            MyChange.FindAndReplace(word, "2222", txtmuherriknomre3.Text);
            MyChange.FindAndReplace(word, "3333", txtbannomre2.Text);
            MyChange.FindAndReplace(word, "4444", txtsassinomre4.Text);
            MyChange.FindAndReplace(word, "6666", txtreng12.Text);

            doc.Save();
        }

        public void reqemler()      //------reqem yazi ile----------------------------------------------------------------
        {
            try { txtherf1.Text = MyChange.ReqemToMetn(Convert.ToDouble(txtlizinqmebleg.Text)); } catch { }
        }

        public void reqemler2()      //------reqem yazi ile---------------------------------------------------------------
        {
            try { txtherf2.Text = MyChange.ReqemToMetn(Convert.ToDouble(txtavadanliqdeyer.Text)); } catch { }
        }

        public void reqemler3()      //------reqem yazi ile---------------------------------------------------------------
        {
            try
            {
                txtherf3.Text = MyChange.ReqemToMetn(Convert.ToDouble(txtavans.Text));
            }
            catch { }
        }

        public void reqemler4()      //------reqem yazi ile---------------------------------------------------------------
        {
            try
            {
                txtherf4.Text = MyChange.ReqemToMetn(Convert.ToDouble(txtmuddet.Text)); 
            }
            catch { }
        }

        public void reqemler5()      //------reqem yazi ile---------------------------------------------------------------
        {
            try
            {
                txtherf5.Text = MyChange.ReqemToMetn(Convert.ToDouble(txtfaiz.Text));
            }
            catch { }
        }

        public void reqemler6()      //------reqem yazi ile---------------------------------------------------------------
        {
            try
            {
                txtherf6.Text = MyChange.ReqemToMetn(Convert.ToDouble(txtbirdefemukafat2.Text));
            }
            catch { }
        }

        public void reqemler7()      //------reqem yazi ile---------------------------------------------------------------
        {
            try
            {
                txtherf7.Text = MyChange.ReqemToMetn(Convert.ToDouble(txtsigorta.Text));
            }
            catch { }
        }

        public void reqemler8()      //------reqem yazi ile---------------------------------------------------------------
        {
            try
            {
                txtherf8.Text = MyChange.ReqemToMetn(Convert.ToDouble(txtnagdilasdirma2.Text));
            }
            catch { }
        }

        public void reqemler9()      //------reqem yazi ile---------------------------------------------------------------
        {
            try
            {
                txtumumimeblegHerf.Text = MyChange.ReqemToMetn(Convert.ToDouble(txtumumimeblegReqem.Text));
            }
            catch { }
        }

        public void reqemler10()      //------reqem yazi ile--------------------------------------------------------------
        {
            try
            {
                txtmukafatHerf.Text = MyChange.ReqemToMetn(Convert.ToDouble(txtmukafatReqem.Text));
            }
            catch { }
        }

        public void reqemler11()      //------reqem yazi ile--------------------------------------------------------------
        {
            try
            {
                txtXerclerHerf.Text = MyChange.ReqemToMetn(Convert.ToDouble(txtXerclerReqem.Text));
            }
            catch {}
        }

        public void reqemler12()      //------reqem yazi ile--------------------------------------------------------------
        {
            try
            {
                txtDebbeFaiziHerf.Text = MyChange.ReqemToMetn(Convert.ToDouble(txtDebbeFaiziReqem.Text));
            }
            catch { }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            SaveRazilasma();

            myrefresLayiheler();
        }

        private void txtlizinqmebleg_TextChanged(object sender, EventArgs e)
        {
            try { txtavans.Text = (Convert.ToDouble(txtavadanliqdeyer.Text) - Convert.ToDouble(txtlizinqmebleg.Text)).ToString();  }
            catch { }

            try { reqemler(); }
            catch { }
        }

        private void txtavadanliqdeyer_TextChanged(object sender, EventArgs e)
        {
            try { txtavans.Text = (Convert.ToDouble(txtavadanliqdeyer.Text) - Convert.ToDouble(txtlizinqmebleg.Text)).ToString(); }
            catch { }

            try { reqemler2(); }
            catch { }
        }

        private void txtavans_TextChanged(object sender, EventArgs e)
        {
            try { reqemler3(); }
            catch { }
        }

        private void txtmuddet_TextChanged(object sender, EventArgs e)
        {
            try { reqemler4(); }
            catch { }
        }

        private void txtfaiz_TextChanged(object sender, EventArgs e)
        {
            try { txtDebbeFaiziReqem.Text = (Convert.ToDouble(txtfaiz.Text)*2).ToString(); }
            catch { }

            try { reqemler5(); }
            catch { }
        }

        private void txtbirdefemukafat_TextChanged(object sender, EventArgs e)
        {
            try
            {
                txtbirdefemukafat2.Text = (Convert.ToDouble(txtbirdefemukafat.Text) * Convert.ToDouble(txtlizinqmebleg.Text) / 100).ToString();
            }
            catch { } 
        }

        private void txtlizinqalan_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    MyData.selectCommand("baza.accdb", "Select * from muqavilelayihe where [Lizinq alan] Like " + "'%" + txtlizinqalan.Text + "%'");
                    MyData.dtmain = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmain);

                    txtlizinqalan.Text=MyData.dtmain.Rows[0]["Lizinq alan"].ToString();
                    txtsatici.Text=MyData.dtmain.Rows[0]["Satıcı"].ToString();
                    txtobyekt.Text=MyData.dtmain.Rows[0]["Lizinq obyekti"].ToString();
                    txtlizinqmebleg.Text=MyData.dtmain.Rows[0]["Lizinq məbləği"].ToString();
                    txtavadanliqdeyer.Text = MyData.dtmain.Rows[0]["Avadanlığın dəyəri"].ToString();
                    txtmuddet.Text = MyData.dtmain.Rows[0]["Lizinqin müddəti"].ToString();
                    txtavans.Text = MyData.dtmain.Rows[0]["Avans"].ToString();
                    txtfaiz.Text = MyData.dtmain.Rows[0]["% dərəcəsi"].ToString();
                    txtqrafik.Text = MyData.dtmain.Rows[0]["Qrafik"].ToString();
                    txtguzest.Text = MyData.dtmain.Rows[0]["Güzəşt müddəti (ay)"].ToString();
                    txtlizmeqsedi.Text = MyData.dtmain.Rows[0]["Lizinqin məqsədi"].ToString();
                    txtsigorta.Text = MyData.dtmain.Rows[0]["Siğorta"].ToString();
                    txtteminat.Text = MyData.dtmain.Rows[0]["Тəminat"].ToString();
                    txtbirdefemukafat.Text = MyData.dtmain.Rows[0]["Birdəfəlik mükafat (%)"].ToString();
                    txtsertler.Text = MyData.dtmain.Rows[0]["Cari şərtlər"].ToString();
                    txtmonitorinq.Text = MyData.dtmain.Rows[0]["İlkin monitorinq"].ToString();
                    txtnagdilasdirma.Text = MyData.dtmain.Rows[0]["Nağdılaşdırma (%)"].ToString();
                    cmbkurator.Text = MyData.dtmain.Rows[0]["Kurator"].ToString();
                    dttarix.Text = MyData.dtmain.Rows[0]["Tarix"].ToString();
                    cmbfhs1.Text = MyData.dtmain.Rows[0]["Müştəri növü"].ToString();
                    cmbfhs2.Text = MyData.dtmain.Rows[0]["Satıcı növü"].ToString();
                }
                catch { }
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            txtlizinqalan.Text = "";
            txtsatici.Text = "";
            txtobyekt.Text = "";
            txtlizinqmebleg.Text = "0";
            txtavadanliqdeyer.Text = "0";
            txtmuddet.Text = "0";
            txtavans.Text = "0";
            txtfaiz.Text = "0";
            txtqrafik.Text = "Annuitet";
            txtguzest.Text = "0";
            txtlizmeqsedi.Text = "";
            txtsigorta.Text = "0";
            txtteminat.Text = "";
            txtbirdefemukafat.Text = "0";
            txtsertler.Text = "";
            txtmonitorinq.Text = "0";
            txtnagdilasdirma.Text = "0";
            cmbkurator.Text = " - ";
            dttarix.Text = "";
            cmbfhs1.Text = "Fiziki şəxs";
            cmbfhs2.Text = "Fiziki şəxs";
           
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox5.Text == "Fiziki şəxs") VoenSatici.Visible = false; else VoenSatici.Visible = true;
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox4.Text == "Fiziki şəxs") voenLizinqalan.Visible = false; else voenLizinqalan.Visible = true;
        }

        private void cmbfhs1_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox4.Text = cmbfhs1.Text;
        }

        private void cmbfhs2_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox5.Text = cmbfhs2.Text;
        }

        private void txtnomre1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    MyData.selectCommand("baza.accdb", "SELECT * FROM etibarnameneqliyyat WHERE c1 like " + "'%" + txtnomre1.Text + "%'");
                }
                catch { };

                MyData.dtmain = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmain);

                try { txtnomre1.Text = MyData.dtmain.Rows[0]["c1"].ToString(); }catch { };
                try { txtbannomre2.Text = MyData.dtmain.Rows[0]["c8"].ToString(); }  catch { };
                try { txtmuherriknomre3.Text = MyData.dtmain.Rows[0]["c9"].ToString(); }catch { };
                try { txtsassinomre4.Text = MyData.dtmain.Rows[0]["c10"].ToString(); } catch { };
                try { txttexpass5.Text = MyData.dtmain.Rows[0]["c12"].ToString(); } catch { };
                try { txtburaxilisili6.Text = MyData.dtmain.Rows[0]["c6"].ToString(); }catch { };
                try { txtpassverilmetarix7.Text = MyData.dtmain.Rows[0]["c7"].ToString(); }catch { };
                try { txtmarkasi8.Text = MyData.dtmain.Rows[0]["c2"].ToString(); } catch { };
                try { txtlizinqalanadi9.Text = MyData.dtmain.Rows[0]["c3"].ToString(); }catch { };
                try { txtlayihe.Text = MyData.dtmain.Rows[0]["c4"].ToString(); }catch { };
                try { txtzavodnomresi11.Text = MyData.dtmain.Rows[0]["c11"].ToString(); }catch { };
                try { txtreng12.Text = MyData.dtmain.Rows[0]["c5"].ToString(); }catch { };
                try { txtQeyd.Text = MyData.dtmain.Rows[0]["c13"].ToString(); }catch { };
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            SaveSaticiRekvizit();

            myrefresRekvizitSatici();
        }

        private void button15_Click(object sender, EventArgs e)
        {
            SaveLizinqalanRekvizit();

            myrefresRekvizit();
        }

        private void textBox34_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    MyData.selectCommand("baza.accdb", "Select * from muqavilerekvizit where [Lizinq alan] Like " + "'%" + textBox34.Text + "%'");
                    MyData.dtmain = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmain);

                    textBox34.Text = MyData.dtmain.Rows[0]["Lizinq alan"].ToString(); //----Satici
                    txtfizsexsiyyetseriya.Text = MyData.dtmain.Rows[0]["Şəxsiyyət vəsiqəsi"].ToString().Substring(0,3);
                    txtfizsexsiyyetnomre.Text = MyData.dtmain.Rows[0]["Şəxsiyyət vəsiqəsi"].ToString().Substring(6, MyData.dtmain.Rows[0]["Şəxsiyyət vəsiqəsi"].ToString().Length-6);
                    txtfizqeydiyyat.Text = MyData.dtmain.Rows[0]["Ş/V qeydiyyat ünvanı"].ToString();
                    txtfizfaktikiunvan.Text = MyData.dtmain.Rows[0]["Ş/V faktiki ünvan"].ToString();
                    txtfizvesiqeveren.Text = MyData.dtmain.Rows[0]["Ş/V verən orqan"].ToString();
                    dtfizvesverilmetarix.Text = MyData.dtmain.Rows[0]["Ş/V verilmə tarixi"].ToString();
                    cmbfiztelkod1.Text = MyData.dtmain.Rows[0]["Əlaqə nömrəsi 1"].ToString().Substring(0,3);
                    txtfiznomre1.Text = MyData.dtmain.Rows[0]["Əlaqə nömrəsi 1"].ToString().Substring(4, MyData.dtmain.Rows[0]["Əlaqə nömrəsi 1"].ToString().Length-4); ;
                    cmbfiztelkod2.Text = MyData.dtmain.Rows[0]["Əlaqə nömrəsi 2"].ToString().Substring(0, 3); ;
                    txtfiznomre2.Text = MyData.dtmain.Rows[0]["Əlaqə nömrəsi 2"].ToString().Substring(4, MyData.dtmain.Rows[0]["Əlaqə nömrəsi 2"].ToString().Length - 4); ; ;
                    cmbfiztelkod3.Text = MyData.dtmain.Rows[0]["Əlaqə nömrəsi 3"].ToString().Substring(0, 3); ;
                    txtfiznomre3.Text = MyData.dtmain.Rows[0]["Əlaqə nömrəsi 3"].ToString().Substring(4, MyData.dtmain.Rows[0]["Əlaqə nömrəsi 3"].ToString().Length - 4); ; ;
                    txtfizrekvizit.Text = MyData.dtmain.Rows[0]["Rekvizitər"].ToString();
                    txtVoen1.Text = MyData.dtmain.Rows[0]["Vöen"].ToString();
                    txtVoenHuquqiUnvan1.Text = MyData.dtmain.Rows[0]["Vöen qeydiyyat ünvanı"].ToString();
                    txtVoenVerenOrqan1.Text = MyData.dtmain.Rows[0]["Vöen verən orqan"].ToString();
                    txtVoenTarix1.Text = MyData.dtmain.Rows[0]["Vöen verilmə tarixi"].ToString();
                    txtDirektor1.Text = MyData.dtmain.Rows[0]["Direktor"].ToString();
                }
                catch { }
            }
        }

        private void txtsatici2_KeyDown(object sender, KeyEventArgs e)
        {
             if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    MyData.selectCommand("baza.accdb", "Select * from muqavilesaticirekvizit where [Satıcı] Like " + "'%" + txtsatici2.Text + "%'");
                    MyData.dtmain = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmain);

                    txtsatici2.Text = MyData.dtmain.Rows[0]["Satıcı"].ToString(); //----Satici
                    textBox21.Text = MyData.dtmain.Rows[0]["Şəxsiyyət vəsiqəsi"].ToString().Substring(0, 3);
                    textBox20.Text = MyData.dtmain.Rows[0]["Şəxsiyyət vəsiqəsi"].ToString().Substring(6, MyData.dtmain.Rows[0]["Şəxsiyyət vəsiqəsi"].ToString().Length - 6);
                    textBox17.Text = MyData.dtmain.Rows[0]["Ş/V qeydiyyat ünvanı"].ToString();
                    textBox18.Text = MyData.dtmain.Rows[0]["Ş/V faktiki ünvan"].ToString();
                    textBox19.Text = MyData.dtmain.Rows[0]["Ş/V verən orqan"].ToString();
                    dateTimePicker2.Text = MyData.dtmain.Rows[0]["Ş/V verilmə tarixi"].ToString();
                    comboBox3.Text = MyData.dtmain.Rows[0]["Əlaqə nömrəsi 1"].ToString().Substring(0,3);
                    maskedTextBox3.Text = MyData.dtmain.Rows[0]["Əlaqə nömrəsi 1"].ToString().Substring(4, MyData.dtmain.Rows[0]["Əlaqə nömrəsi 1"].ToString().Length-4); ;
                    comboBox2.Text = MyData.dtmain.Rows[0]["Əlaqə nömrəsi 2"].ToString().Substring(0, 3); ;
                    maskedTextBox2.Text = MyData.dtmain.Rows[0]["Əlaqə nömrəsi 2"].ToString().Substring(4, MyData.dtmain.Rows[0]["Əlaqə nömrəsi 2"].ToString().Length - 4); ; ;
                    comboBox1.Text = MyData.dtmain.Rows[0]["Əlaqə nömrəsi 3"].ToString().Substring(0, 3); ;
                    maskedTextBox1.Text = MyData.dtmain.Rows[0]["Əlaqə nömrəsi 3"].ToString().Substring(4, MyData.dtmain.Rows[0]["Əlaqə nömrəsi 3"].ToString().Length - 4); ; ;
                    richTextBox1.Text = MyData.dtmain.Rows[0]["Rekvizitər"].ToString();
                    txtVoen2.Text = MyData.dtmain.Rows[0]["Vöen"].ToString();
                    txtVoenHuquqiUnvan2.Text = MyData.dtmain.Rows[0]["Vöen qeydiyyat ünvanı"].ToString();
                    txtVoenVerenOrqan2.Text = MyData.dtmain.Rows[0]["Vöen verən orqan"].ToString();
                    txtVoenTarix2.Text = MyData.dtmain.Rows[0]["Vöen verilmə tarixi"].ToString();
                    txtDirektor2.Text = MyData.dtmain.Rows[0]["Direktor"].ToString();

                }
                catch { }
            }
        
        }

        private void Yeni_Muqavile_Load(object sender, EventArgs e)
        {
            myrefresLayiheler();
            myrefresRekvizit();
            myrefresRekvizitSatici();
        }

        private void txtsigorta_TextChanged(object sender, EventArgs e)
        {
            try { reqemler7(); }
            catch { }
        }

        private void txtnagdilasdirma_TextChanged(object sender, EventArgs e)
        {
            try { txtnagdilasdirma2.Text = (Convert.ToDouble(txtavadanliqdeyer.Text) * Convert.ToDouble(txtnagdilasdirma.Text) / 100).ToString(); }
            catch { }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            PrintDaxiliMaliyye();
        }

        private void txtobyekt_TextChanged(object sender, EventArgs e)
        {
            txtteminat.Text = txtobyekt.Text;
        }

        private void txtumumimeblegReqem_TextChanged(object sender, EventArgs e)
        {
            try { reqemler9(); }
            catch { }
        }

        private void txtmukafatReqem_TextChanged(object sender, EventArgs e)
        {
            try { reqemler10(); }
            catch { }
        }

        private void txtXerclerReqem_TextChanged(object sender, EventArgs e)
        {
            try { reqemler11(); }
            catch { }
        }

        private void txtnagdilasdirma2_TextChanged(object sender, EventArgs e)
        {
            try { reqemler8(); }
            catch { }
        }

        private void txtbirdefemukafat2_TextChanged(object sender, EventArgs e)
        {
            try { reqemler6(); }
            catch { }
        }

        private void txtDebbeFaiziReqem_TextChanged(object sender, EventArgs e)
        {
            try { reqemler12(); }
            catch { }
        }

        private void autoDoldurToolStripMenuItem_Click(object sender, EventArgs e)
        {
            txtfizrekvizit.Text = textBox34.Text + Environment.NewLine + "Ş/V: " + txtfizsexsiyyetseriya.Text + " № " + txtfizsexsiyyetnomre.Text + " (" + txtfizvesiqeveren.Text + ", " + dtfizvesverilmetarix.Text + ")" + Environment.NewLine + "Qeydiyyat ünvanı: " + txtfizqeydiyyat.Text + Environment.NewLine + "Faktiki ünvanı: " + txtfizfaktikiunvan.Text;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            PrintElave1();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            PrintYekunSifaris();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            PrintQrafik();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            PrintAlqiSatqi();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            PrintAlqiSatqiElave1();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            PrintQebulİstifade();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            SaveNeqliyyat();
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            richTextBox1.Text = txtsatici2.Text + Environment.NewLine + "Ş/V: " + textBox21.Text + " № " + textBox20.Text + " (" + textBox19.Text + ", " + dateTimePicker2.Text + ")" + Environment.NewLine + "Qeydiyyat ünvanı: " + textBox17.Text + Environment.NewLine + "Faktiki ünvanı: " + textBox18.Text;
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                MyData.selectCommand("baza.accdb", "Select * from muqavilelayihe where [Lizinq alan] Like '%" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Lizinq alan"].Value.ToString() + "%'");
                MyData.dtmain = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmain);

                txtlizinqalan.Text = MyData.dtmain.Rows[0]["Lizinq alan"].ToString();
                txtsatici.Text = MyData.dtmain.Rows[0]["Satıcı"].ToString();
                txtobyekt.Text = MyData.dtmain.Rows[0]["Lizinq obyekti"].ToString();
                txtlizinqmebleg.Text = MyData.dtmain.Rows[0]["Lizinq məbləği"].ToString();
                txtavadanliqdeyer.Text = MyData.dtmain.Rows[0]["Avadanlığın dəyəri"].ToString();
                txtmuddet.Text = MyData.dtmain.Rows[0]["Lizinqin müddəti"].ToString();
                txtavans.Text = MyData.dtmain.Rows[0]["Avans"].ToString();
                txtfaiz.Text = MyData.dtmain.Rows[0]["% dərəcəsi"].ToString();
                txtqrafik.Text = MyData.dtmain.Rows[0]["Qrafik"].ToString();
                txtguzest.Text = MyData.dtmain.Rows[0]["Güzəşt müddəti (ay)"].ToString();
                txtlizmeqsedi.Text = MyData.dtmain.Rows[0]["Lizinqin məqsədi"].ToString();
                txtsigorta.Text = MyData.dtmain.Rows[0]["Siğorta"].ToString();
                txtteminat.Text = MyData.dtmain.Rows[0]["Тəminat"].ToString();
                txtbirdefemukafat.Text = MyData.dtmain.Rows[0]["Birdəfəlik mükafat (%)"].ToString();
                txtsertler.Text = MyData.dtmain.Rows[0]["Cari şərtlər"].ToString();
                txtmonitorinq.Text = MyData.dtmain.Rows[0]["İlkin monitorinq"].ToString();
                txtnagdilasdirma.Text = MyData.dtmain.Rows[0]["Nağdılaşdırma (%)"].ToString();
                cmbkurator.Text = MyData.dtmain.Rows[0]["Kurator"].ToString();
                dttarix.Text = MyData.dtmain.Rows[0]["Tarix"].ToString();
                cmbfhs1.Text = MyData.dtmain.Rows[0]["Müştəri növü"].ToString();
                cmbfhs2.Text = MyData.dtmain.Rows[0]["Satıcı növü"].ToString();

                tabControl1.SelectedTab = tabPage1;
            }
            catch { }
        }

        private void dataGridView2_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                MyData.selectCommand("baza.accdb", "Select * from muqavilerekvizit where [Lizinq alan] Like " + "'%" + dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells["Lizinq alan"].Value.ToString() + "%'");
                MyData.dtmain = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmain);

                textBox34.Text = MyData.dtmain.Rows[0]["Lizinq alan"].ToString(); //----Satici
                txtfizsexsiyyetseriya.Text = MyData.dtmain.Rows[0]["Şəxsiyyət vəsiqəsi"].ToString().Substring(0, 3);
                txtfizsexsiyyetnomre.Text = MyData.dtmain.Rows[0]["Şəxsiyyət vəsiqəsi"].ToString().Substring(6, MyData.dtmain.Rows[0]["Şəxsiyyət vəsiqəsi"].ToString().Length - 6);
                txtfizqeydiyyat.Text = MyData.dtmain.Rows[0]["Ş/V qeydiyyat ünvanı"].ToString();
                txtfizfaktikiunvan.Text = MyData.dtmain.Rows[0]["Ş/V faktiki ünvan"].ToString();
                txtfizvesiqeveren.Text = MyData.dtmain.Rows[0]["Ş/V verən orqan"].ToString();
                dtfizvesverilmetarix.Text = MyData.dtmain.Rows[0]["Ş/V verilmə tarixi"].ToString();
                cmbfiztelkod1.Text = MyData.dtmain.Rows[0]["Əlaqə nömrəsi 1"].ToString().Substring(0, 3);
                txtfiznomre1.Text = MyData.dtmain.Rows[0]["Əlaqə nömrəsi 1"].ToString().Substring(4, MyData.dtmain.Rows[0]["Əlaqə nömrəsi 1"].ToString().Length - 4); ;
                cmbfiztelkod2.Text = MyData.dtmain.Rows[0]["Əlaqə nömrəsi 2"].ToString().Substring(0, 3); ;
                txtfiznomre2.Text = MyData.dtmain.Rows[0]["Əlaqə nömrəsi 2"].ToString().Substring(4, MyData.dtmain.Rows[0]["Əlaqə nömrəsi 2"].ToString().Length - 4); ; ;
                cmbfiztelkod3.Text = MyData.dtmain.Rows[0]["Əlaqə nömrəsi 3"].ToString().Substring(0, 3); ;
                txtfiznomre3.Text = MyData.dtmain.Rows[0]["Əlaqə nömrəsi 3"].ToString().Substring(4, MyData.dtmain.Rows[0]["Əlaqə nömrəsi 3"].ToString().Length - 4); ; ;
                txtfizrekvizit.Text = MyData.dtmain.Rows[0]["Rekvizitər"].ToString();
                txtVoen1.Text = MyData.dtmain.Rows[0]["Vöen"].ToString();
                txtVoenHuquqiUnvan1.Text = MyData.dtmain.Rows[0]["Vöen qeydiyyat ünvanı"].ToString();
                txtVoenVerenOrqan1.Text = MyData.dtmain.Rows[0]["Vöen verən orqan"].ToString();
                txtVoenTarix1.Text = MyData.dtmain.Rows[0]["Vöen verilmə tarixi"].ToString();
                txtDirektor1.Text = MyData.dtmain.Rows[0]["Direktor"].ToString();

                tabControl1.SelectedTab = tabPage2;
            }
            catch { }
        }

        private void dataGridView3_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                MyData.selectCommand("baza.accdb", "Select * from muqavilesaticirekvizit where [Satıcı] Like " + "'%" + dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells["Satıcı"].Value.ToString() + "%'");
                MyData.dtmain = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmain);

                txtsatici2.Text = MyData.dtmain.Rows[0]["Satıcı"].ToString(); //----Satici
                textBox21.Text = MyData.dtmain.Rows[0]["Şəxsiyyət vəsiqəsi"].ToString().Substring(0, 3);
                textBox20.Text = MyData.dtmain.Rows[0]["Şəxsiyyət vəsiqəsi"].ToString().Substring(6, MyData.dtmain.Rows[0]["Şəxsiyyət vəsiqəsi"].ToString().Length - 6);
                textBox17.Text = MyData.dtmain.Rows[0]["Ş/V qeydiyyat ünvanı"].ToString();
                textBox18.Text = MyData.dtmain.Rows[0]["Ş/V faktiki ünvan"].ToString();
                textBox19.Text = MyData.dtmain.Rows[0]["Ş/V verən orqan"].ToString();
                dateTimePicker2.Text = MyData.dtmain.Rows[0]["Ş/V verilmə tarixi"].ToString();
                comboBox3.Text = MyData.dtmain.Rows[0]["Əlaqə nömrəsi 1"].ToString().Substring(0, 3);
                maskedTextBox3.Text = MyData.dtmain.Rows[0]["Əlaqə nömrəsi 1"].ToString().Substring(4, MyData.dtmain.Rows[0]["Əlaqə nömrəsi 1"].ToString().Length - 4); ;
                comboBox2.Text = MyData.dtmain.Rows[0]["Əlaqə nömrəsi 2"].ToString().Substring(0, 3); ;
                maskedTextBox2.Text = MyData.dtmain.Rows[0]["Əlaqə nömrəsi 2"].ToString().Substring(4, MyData.dtmain.Rows[0]["Əlaqə nömrəsi 2"].ToString().Length - 4); ; ;
                comboBox1.Text = MyData.dtmain.Rows[0]["Əlaqə nömrəsi 3"].ToString().Substring(0, 3); ;
                maskedTextBox1.Text = MyData.dtmain.Rows[0]["Əlaqə nömrəsi 3"].ToString().Substring(4, MyData.dtmain.Rows[0]["Əlaqə nömrəsi 3"].ToString().Length - 4); ; ;
                richTextBox1.Text = MyData.dtmain.Rows[0]["Rekvizitər"].ToString();
                txtVoen2.Text = MyData.dtmain.Rows[0]["Vöen"].ToString();
                txtVoenHuquqiUnvan2.Text = MyData.dtmain.Rows[0]["Vöen qeydiyyat ünvanı"].ToString();
                txtVoenVerenOrqan2.Text = MyData.dtmain.Rows[0]["Vöen verən orqan"].ToString();
                txtVoenTarix2.Text = MyData.dtmain.Rows[0]["Vöen verilmə tarixi"].ToString();
                txtDirektor2.Text = MyData.dtmain.Rows[0]["Direktor"].ToString();

                tabControl1.SelectedTab = tabPage5;

            }
            catch { }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            PrintDaxiliMaliyye();
            PrintElave1();
            PrintYekunSifaris();
            PrintQrafik();
            PrintAlqiSatqi();
            PrintAlqiSatqiElave1();
            PrintQebulİstifade();
        }

    }
}
