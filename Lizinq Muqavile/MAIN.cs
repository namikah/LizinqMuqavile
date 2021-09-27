using Nsoft;
using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace Lizinq_Muqavile
{
    public partial class Muqavile : Form
    {
        public Muqavile()
        {
            InitializeComponent();
        }

        //Excel.Application oXL;
        //Excel._Workbook oWB;
        //Excel._Worksheet oSheet;
        //Excel._Worksheet bSheet;
        //Excel._Worksheet cSheet;

        private void BuGuneQeydler()
        {
            DateTime dt = DateTime.Now;

            MyData.selectCommand("baza.accdb", "Select * from Qeydler Where c1 Like '%" + dt.AddDays(1).ToShortDateString() + "%' or c1 Like '%" + dt.ToShortDateString() + "%' or c1 Like '%" + dt.AddDays(-1).ToShortDateString() + "%' or c1 Like '%" + dt.AddDays(-2).ToShortDateString() + "%'");
            MyData.dtmainQeydler = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainQeydler);

            lbXeberler.Text = "GÜNÜN MƏZƏNNƏSİ: 1 USD - " + MyChange.Mezenne("USD") + ", 1 EUR - " + MyChange.Mezenne("EUR") + ", 1 RUB - " + MyChange.Mezenne("RUB") + ", 1 TRY - " + MyChange.Mezenne("TRY") + " (https://www.cbar.az/)";

            if (MyData.dtmainQeydler.Rows.Count != 0)
            {
                lbXeberler.Text += " ↔ QEYDLƏR: ";
                for (int i = 0; i < MyData.dtmainQeydler.Rows.Count; i++)
                {
                    lbXeberler.Text += MyData.dtmainQeydler.Rows[i][1].ToString().Substring(0, 10) + " " + MyData.dtmainQeydler.Rows[i][2].ToString() + " ↔ ";
                }
            }

           

        }

        public void BuGuneMezenne()
        {
                kursToolStripMenuItem.Text = "Məzənnə (1 USD - " + MyChange.Mezenne("USD") + ")";
        }

        private void Form1_Load(object sender, EventArgs e)
        {
                if (!MyCheck.LisenziyaYoxla())
                {
                    MessageBox.Show("Lisenziyanız Yoxdur!");

                    Lisenziya lisenziya = new Lisenziya();
                    lisenziya.ShowDialog();

                    if (!MyCheck.LisenziyaYoxla())
                    {
                        Environment.Exit(1);
                        Application.Exit();
                    }
                }

            MainLogo mainlogo = new MainLogo();
            mainlogo.ShowDialog();

            BuGuneMezenne();
            BuGuneQeydler();

            base.Text = Environment.UserName;

            timer1.Enabled = true;

            MyChange.SetKeyboardLayout(MyChange.GetInputLanguageByName("AZ"));
        }
  
      
        private void timer1_Tick(object sender, EventArgs e)
        {
            lbXeberler.Left -= 2;
            if (lbXeberler.Right < 0) lbXeberler.Left = base.Width;
        }

        private void Muqavile_FormClosed(object sender, FormClosedEventArgs e)
        {
            DateTime dt = DateTime.Now;

            try { File.Copy(@"baza.accdb", @"BACKUP\\" + dt.Day + " baza.accdb", true); }
            catch { MessageBox.Show("Backup alınmadı."); }
        }


        private void musiqiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (musiqiToolStripMenuItem.ForeColor == Color.Black) { axWMPlayer1.URL = "main.mp3"; musiqiToolStripMenuItem.ForeColor = Color.Red; return; }
            if (musiqiToolStripMenuItem.ForeColor == Color.Red) { musiqiToolStripMenuItem.ForeColor = Color.Black; axWMPlayer1.close(); return; }
        }

        private void hesabMaşınıToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("calc.exe");
        }

        private void emailToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sekillerFalse();

            System.Diagnostics.Process.Start("outlook.exe");

            sekillerTrue();
        }


        private void kursToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sekillerFalse();

            DollarKurs.Form1 form1 = new DollarKurs.Form1();
            form1.ShowDialog();

            sekillerTrue();
        }

        private void sekillerTrue()
        {
            lbXeberler.Visible = true;
            lbXeberler.Left = base.Width;
            sekil1.Visible = true;
            knopka1.Visible = true;
            sekil2.Visible = true;
            knopka2.Visible = true;
            sekil3.Visible = true;
            knopka3.Visible = true;
            sekil4.Visible = true;
            knopka4.Visible = true;
            sekil5.Visible = true;
            knopka5.Visible = true;
            sekil6.Visible = true;
            knopka6.Visible = true;
            sekil7.Visible = true;
            knopka7.Visible = true;
            sekil8.Visible = true;
            knopka8.Visible = true;
            sekil9.Visible = true;
            knopka9.Visible = true;
            sekil10.Visible = true;
            knopka10.Visible = true;
            sekil11.Visible = true;
            knopka11.Visible = true;
            sekil12.Visible = true;
            knopka12.Visible = true;
            sekil13.Visible = true;
            knopka13.Visible = true;
            sekil14.Visible = true;
            knopka14.Visible = true;
            sekil15.Visible = true;
            knopka15.Visible = true;
            sekil16.Visible = true;
            knopka16.Visible = true;
            sekil17.Visible = true;
            knopka17.Visible = true;
            sekil18.Visible = true;
            knopka18.Visible = true;
        }

        private void sekillerFalse()
        {
            lbXeberler.Visible = false;
            sekil1.Visible = false;
            knopka1.Visible = false;
            sekil2.Visible = false;
            knopka2.Visible = false;
            sekil3.Visible = false;
            knopka3.Visible = false;
            sekil4.Visible = false;
            knopka4.Visible = false;
            sekil5.Visible = false;
            knopka5.Visible = false;
            sekil6.Visible = false;
            knopka6.Visible = false;
            sekil7.Visible = false;
            knopka7.Visible = false;
            sekil8.Visible = false;
            knopka8.Visible = false;
            sekil9.Visible = false;
            knopka9.Visible = false;
            sekil10.Visible = false;
            knopka10.Visible = false;
            sekil11.Visible = false;
            knopka11.Visible = false;
            sekil12.Visible = false;
            knopka12.Visible = false;
            sekil13.Visible = false;
            knopka13.Visible = false;
            sekil14.Visible = false;
            knopka14.Visible = false;
            sekil15.Visible = false;
            knopka15.Visible = false;
            sekil16.Visible = false;
            knopka16.Visible = false;
            sekil17.Visible = false;
            knopka17.Visible = false;
            sekil18.Visible = false;
            knopka18.Visible = false;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            MyCheck.ParolYoxla();
            if (!MyCheck.Parolicaze) return;

            sekillerFalse();

            Telefon Telefon = new Telefon();
            Telefon.Show();

            sekillerTrue();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            MyCheck.ParolYoxla();
            if (!MyCheck.Parolicaze) return;

            sekillerFalse();

            Odenisler odenisler = new Odenisler();
            odenisler.Show();

            sekillerTrue();
        }


        private void sekil2_Click(object sender, EventArgs e)
        {
            MyCheck.ParolYoxla();
            if (!MyCheck.Parolicaze) return;

            sekillerFalse();

            EtibarnameEsas EtibarnameEsas = new EtibarnameEsas();
            EtibarnameEsas.Show();

            sekillerTrue();
        }

        private void sekil3_Click(object sender, EventArgs e)
        {
            MyCheck.ParolYoxla();
            if (!MyCheck.Parolicaze) return;

            sekillerFalse();

            Kassa kassa1 = new Kassa();
            kassa1.Show();

            sekillerTrue();
        }

        private void sekil4_Click(object sender, EventArgs e)
        {
            MyCheck.ParolYoxla();
            if (!MyCheck.Parolicaze) return;

            sekillerFalse();

            Plateshka Plateshka = new Plateshka();
            Plateshka.Show();

            sekillerTrue();
        }

        private void sekil5_Click(object sender, EventArgs e)
        {
            MyCheck.ParolYoxla();
            if (!MyCheck.Parolicaze) return;

            sekillerFalse();

            LizinqCalculator LizinqCalculator = new LizinqCalculator();
            LizinqCalculator.Show();

            sekillerTrue();
        }

        private void sekil8_Click(object sender, EventArgs e)
        {
            MyCheck.ParolYoxla();
            if (!MyCheck.Parolicaze) return;

            sekillerFalse();

            DollarKurs.Form1 form1 = new DollarKurs.Form1();
            form1.ShowDialog();

            sekillerTrue();
        }


        private void knopka2_Click(object sender, EventArgs e)
        {
            sekil2_Click(sender, e);
        }

        private void knopka3_Click(object sender, EventArgs e)
        {
            sekil3_Click(sender, e);
        }

        private void knopka4_Click(object sender, EventArgs e)
        {
            sekil4_Click(sender, e);
        }

        private void knopka5_Click(object sender, EventArgs e)
        {
            sekil5_Click(sender, e);
        }

        private void knopka6_Click(object sender, EventArgs e)
        {
            button12_Click(sender, e);
        }

        private void knopka7_Click(object sender, EventArgs e)
        {
            button11_Click(sender, e);
        }

        private void knopka8_Click(object sender, EventArgs e)
        {
            sekil8_Click(sender, e);
        }

        private void lizinqKalkulyatorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sekil5_Click(sender, e);
        }

        private void scanToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("WFS.exe");
        }

        private void sekil9_Click(object sender, EventArgs e)
        {
            MyCheck.ParolYoxla();
            if (!MyCheck.Parolicaze) return;

            sekillerFalse();

            Ohdelik_Verilmesi ohdelik = new Ohdelik_Verilmesi();
            ohdelik.Show();

            sekillerTrue();
        }

        private void knopka9_Click(object sender, EventArgs e)
        {
            sekil9_Click(sender, e);
        }

        private void sekil10_Click(object sender, EventArgs e)
        {
            MyCheck.ParolYoxla();
            if (!MyCheck.Parolicaze) return;

            sekillerFalse();

            DYP qayi = new DYP();
            qayi.Show();

            sekillerTrue();
        }

        private void knopka10_Click(object sender, EventArgs e)
        {
            sekil10_Click(sender, e);
        }

        private void sekil11_Click(object sender, EventArgs e)
        {
            MyCheck.ParolYoxla();
            if (!MyCheck.Parolicaze) return;

            sekillerFalse();

            Mulkiyyete_Verme mulkiyyeteverme = new Mulkiyyete_Verme();
            mulkiyyeteverme.Show();

            sekillerTrue();
        }

        private void sekil12_Click(object sender, EventArgs e)
        {
            MyCheck.ParolYoxla();
            if (!MyCheck.Parolicaze) return;

            sekillerFalse();

            Emeliyyatlar elaveler = new Emeliyyatlar();
            elaveler.Show();

            sekillerTrue();
        }

        private void knopka12_Click(object sender, EventArgs e)
        {
            sekil12_Click(sender,e);
        }

        private void mMXToolStripMenuItem_Click(object sender, EventArgs e)
        {
            button14_Click(sender, e);
        }

        private void əlaqəToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Elaqe elaqe = new Elaqe();
            elaqe.Show();
        }

        private void köməkToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Mail:" + Environment.NewLine + "rashadim@agleasing.az" + Environment.NewLine + "namikah@agleasing.az" + Environment.NewLine + Environment.NewLine + "Tel: (012) 4975017 (ext:1702, ext:1703)", "Kömək");

        }

        private void knopka11_Click(object sender, EventArgs e)
        {
            sekil11_Click(sender, e);
        }

        private void ödənişQəbziToolStripMenuItem_Click(object sender, EventArgs e)
        {
            button12_Click_1(sender, e);
        }

        private void button12_Click_1(object sender, EventArgs e)
        {
            MyCheck.ParolYoxla();
            if (!MyCheck.Parolicaze) return;

            sekillerFalse();

            Qəbz Qəbz = new Qəbz();
            Qəbz.Show();

            sekillerTrue();
        }

        private void button11_Click_1(object sender, EventArgs e)
        {
            button12_Click_1(sender, e);
        }

        private void button14_Click(object sender, EventArgs e)
        {
            MyCheck.ParolYoxla();
            if (!MyCheck.Parolicaze) return;

            sekillerFalse();

            MMX MMX = new MMX();
            MMX.Show();

            sekillerTrue();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            button14_Click(sender, e);
        }

        private void knopka15_Click(object sender, EventArgs e)
        {
            sekil15_Click(sender, e);
        }

        private void sekil15_Click(object sender, EventArgs e)
        {
            MyCheck.ParolYoxla();
            if (!MyCheck.Parolicaze) return;

            sekillerFalse();

            Qeydler qeydler = new Qeydler();
            qeydler.Show();

            sekillerTrue();
        }

        private void tehvilTeslimToolStripMenuItem_Click(object sender, EventArgs e)
        {
            button12_Click_2(sender, e);
        }

        private void portfelToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Portfel portfel = new Portfel();
            portfel.Show();
        }

        private void label124_Click_1(object sender, EventArgs e)
        {
            Qeydler qeydler = new Qeydler();
            qeydler.Show();
        }

        private void sekil1_Click(object sender, EventArgs e)
        {
            MyCheck.ParolYoxla();
            if (!MyCheck.Parolicaze) return;

            sekillerFalse();

            Yeni_Muqavile yenimuqavile = new Yeni_Muqavile();
            yenimuqavile.Show();

            sekillerTrue();
        }

        private void knopka1_Click(object sender, EventArgs e)
        {
            sekil1_Click(sender, e);
        }

        private void rekvizitlərToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Rekvizitler rekvizitler = new Rekvizitler();
            rekvizitler.Show();
        }

        private void qeydlərToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            Qeydler qeydler = new Qeydler();
            qeydler.Show();
        }

        private void bildirişToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sekil18_Click(sender, e);
        }

        private void dYPToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sekil10_Click(sender, e);
        }

        private void DaxiliMaliyyeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sekil1_Click(sender, e);
        }

        private void öhdəliklərinVerilməsiToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            sekil9_Click(sender, e);
        }

        private void mülkiyyətəVerməToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sekil11_Click(sender, e);
        }

        private void şəkillərToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Sekiller sekiller = new Sekiller();
            sekiller.Show();
        }

        private void button12_Click_2(object sender, EventArgs e)
        {
            MyCheck.ParolYoxla();
            if (!MyCheck.Parolicaze) return;

            sekillerFalse();

            TehvilTeslim TehvilTeslim = new TehvilTeslim();
            TehvilTeslim.Show();

            sekillerTrue();
        }

        private void button11_Click_2(object sender, EventArgs e)
        {
            button12_Click_2(sender, e);
        }

        private void button10_Click(object sender, EventArgs e)
        {
            MyCheck.ParolYoxla();
            if (!MyCheck.Parolicaze) return;

            sekillerFalse();

            İcbari icbari = new İcbari();
            icbari.Show();

            sekillerTrue();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            button10_Click(sender, e);
        }

        private void tsmcixis_Click(object sender, EventArgs e)
        {
            base.Close();
        }

        private void etibarnaməToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sekil2_Click(sender, e);
        }

        private void sekil18_Click(object sender, EventArgs e)
        {
            MyCheck.ParolYoxla();
            if (!MyCheck.Parolicaze) return;

            sekillerFalse();

            Bildiris bildiris = new Bildiris();
            bildiris.Show();

            sekillerTrue();
        }

        private void knopka18_Click(object sender, EventArgs e)
        {
            sekil18_Click(sender, e);
        }

        private void havaHaqqındaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Hava hava = new Hava();
            hava.Show();

        }

        private void kreditKalkulyatorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Qrafik qrafik = new Qrafik();
            qrafik.Show();
        }

        private void işçilərToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Isciler isciler = new Isciler();
            isciler.Show();
        }

        private void mektublarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Mektublar mektublar = new Mektublar();
            mektublar.Show();
        }

        private void licenziyaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Lisenziya lisenziya = new Lisenziya();
            lisenziya.Show();
        }

        private void KassaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sekil3_Click(sender, e);
        }

        private void ÖdənişTapşırığıToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            sekil4_Click(sender, e);
        }

        private void ÖdənişlərToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            button11_Click(sender, e);
        }

        private void IcbariSiğortaToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            button10_Click(sender, e);
        }

        private void HaqqındaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("NSoft - 2021","info");
        }

        private void AdminPanelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MyCheck.ParolYoxla();
            if (!MyCheck.Parolicaze) return;

            if (!MyCheck.ParolAdminYesNo()) return;

            AdminPanel adminpanel = new AdminPanel();
            adminpanel.Show();
        }

        private void Tsmprint_Click(object sender, EventArgs e)
        {
            Odenisler odenisler = new Odenisler();
            odenisler.wordToolStripMenuItem_Click(sender,e);
        }

        private void HaqqındaToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Haqqinda haqqinda = new Haqqinda();
            haqqinda.ShowDialog();
        }

        private void ParolToolStripMenuItem_Click(object sender, EventArgs e)
        {

            YeniParol YeniParol = new YeniParol();
            YeniParol.ShowDialog();
        }
    }
}
