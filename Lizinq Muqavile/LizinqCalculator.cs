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

namespace Lizinq_Muqavile
{
    public partial class LizinqCalculator : Form
    {

        Excel.Application oXL;
        Excel._Workbook oWB;
        Excel._Worksheet oSheet;

        public LizinqCalculator()
        {
            InitializeComponent();
        }

        private void PrintXercler()    //-----------------Print rasxod ucun-------------------------------
        {
            try
            {
                File.Copy("Bos.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Xercler.xlsx", true);
            }
            catch { MessageBox.Show("Bos.xlsx tapılmadı."); }

            int a = 5;

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Xercler.xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            // oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

            oSheet.Cells[2, 1] = "AGLizinq QSC";
            oSheet.Cells[3, 1] = "MK: 138023";

            if (label20.Text != "0 AZN") { oSheet.Cells[a, 1] = textBox1.Text + " Komissiya - " + label20.Text; a = a + 1; }
            if (label21.Text != "0 AZN") { oSheet.Cells[a, 1] = textBox1.Text + " Nağdılaşdırma - " + label21.Text; a = a + 1; }
            if (label23.Text != "0 AZN") { oSheet.Cells[a, 1] = textBox1.Text + " Etibarnamə - " + label23.Text; a = a + 1; }
            if (label24.Text != "0 AZN") { oSheet.Cells[a, 1] = textBox1.Text + " Siğorta - " + label24.Text; a = a + 1; }
            if (label17.Text != "0 AZN") { oSheet.Cells[a, 1] = textBox1.Text + " Avans - " + label17.Text; a = a + 1; }
            if (label22.Text != "0 AZN") { oSheet.Cells[a, 1] = textBox1.Text + " Köçürmə - " + label22.Text; a = a + 1; }

            oSheet.Rows.Font.Bold = true;
            oSheet.Rows.Font.Size = 14;
            oSheet.Columns.AutoFit();
            oSheet.Rows.AutoFit();
            oSheet.PrintOut();
            oWB.Close(SaveChanges: false);
            oXL.Workbooks.Close();


        }

        private void txtavadanliqdeyer_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    txtavans.Text = (100 - Convert.ToDouble(txtmalyelesme.Text) * 100 / Convert.ToDouble(txtavadanliqdeyer.Text)).ToString();
                }
                catch { };
            }
        }

        private void txtmalyelesme_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    txtavans.Text = (100 - Convert.ToDouble(txtmalyelesme.Text) * 100 / Convert.ToDouble(txtavadanliqdeyer.Text)).ToString();
                }
                catch { };
            }
        }

        private void txtavans_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    txtmalyelesme.Text = (Convert.ToDouble(txtavadanliqdeyer.Text) * (100 - Convert.ToDouble(txtavans.Text)) / 100).ToString();
                }
                catch { };
            }
        }


        private void button1_Click(object sender, EventArgs e)
        {
            int k1 = 0;
            string k2 = "";

            if (txtavadanliqdeyer.Text == "" || txtavadanliqdeyer.Text == "0" || Convert.ToDouble(txtmalyelesme.Text) > Convert.ToDouble(txtavadanliqdeyer.Text)) txtavadanliqdeyer.Text = txtmalyelesme.Text;
            if (txtsigortameblegi.Text == "" || txtsigortameblegi.Text == "0" || Convert.ToDouble(txtsigortameblegi.Text) < Convert.ToDouble(txtavadanliqdeyer.Text)) txtsigortameblegi.Text = txtavadanliqdeyer.Text;

            try
            {
                txtavans.Text = (100 - Convert.ToDouble(txtmalyelesme.Text) * 100 / Convert.ToDouble(txtavadanliqdeyer.Text)).ToString();
            }
            catch { };

            lbayliqodenis.Text = (Convert.ToDouble(txtmalyelesme.Text) * (Convert.ToDouble(txtfaiz.Text) / 100 / 12) / (1 - 1 / Math.Pow((1 + Convert.ToDouble(txtfaiz.Text) / 100 / 12), Convert.ToDouble(txtmuddet.Text)))).ToString();
            label15.Text = "Komissiya: " + Convert.ToDouble(txtmalyelesme.Text) * Convert.ToDouble(txtkomissiya.Text) / 100 + " AZN";
            label19.Text = "Avans: " + (Convert.ToDouble(txtavadanliqdeyer.Text) - Convert.ToDouble(txtmalyelesme.Text)).ToString() + " AZN";
            label16.Text = "Nağdılaşdırma: " + (Convert.ToDouble(txtavadanliqdeyer.Text) * Convert.ToDouble(txtnagdilasdirma.Text) / 100).ToString() + " AZN";
            lbkocurme.Text = "Köçürmə: " + (Convert.ToDouble(txtavadanliqdeyer.Text) * Convert.ToDouble(txtkocurme.Text) / 100).ToString() + " AZN";
            lbDYP.Text = "DYP: " + txtqeyriresmi.Text + " AZN";
            label8.Text = "Etibarnamə: " + Convert.ToDouble(txtetibarname.Text) * 3 + " AZN";

            if (txtsigortanovu.Text == "Minik avtomobilləri" && txtsigortamuddeti.Text == "12") lbsigorta.Text = "Siğorta: " + (Convert.ToDouble(txtsigortameblegi.Text) * 0.0315).ToString() + " AZN";
            else if (txtsigortanovu.Text == "Minik avtomobilləri" && txtsigortamuddeti.Text == "18") lbsigorta.Text = "Siğorta: " + (Convert.ToDouble(txtsigortameblegi.Text) * 0.0365).ToString() + " AZN";
            else if (txtsigortanovu.Text == "Minik avtomobilləri" && txtsigortamuddeti.Text == "24") lbsigorta.Text = "Siğorta: " + (Convert.ToDouble(txtsigortameblegi.Text) * 0.0475).ToString() + " AZN";
            else if (txtsigortanovu.Text == "Minik avtomobilləri" && txtsigortamuddeti.Text == "30") lbsigorta.Text = "Siğorta: " + (Convert.ToDouble(txtsigortameblegi.Text) * 0.0555).ToString() + " AZN";
            else if (txtsigortanovu.Text == "Minik avtomobilləri" && txtsigortamuddeti.Text == "36") lbsigorta.Text = "Siğorta: " + (Convert.ToDouble(txtsigortameblegi.Text) * 0.0610).ToString() + " AZN";
            else if (txtsigortanovu.Text == "Minik avtomobilləri" && txtsigortamuddeti.Text == "42") lbsigorta.Text = "Siğorta: " + (Convert.ToDouble(txtsigortameblegi.Text) * 0.0685).ToString() + " AZN";
            else if (txtsigortanovu.Text == "Minik avtomobilləri" && txtsigortamuddeti.Text == "48") lbsigorta.Text = "Siğorta: " + (Convert.ToDouble(txtsigortameblegi.Text) * 0.0715).ToString() + " AZN";
            else if (txtsigortanovu.Text == "Minik avtomobilləri" && txtsigortamuddeti.Text == "54") lbsigorta.Text = "Siğorta: " + (Convert.ToDouble(txtsigortameblegi.Text) * 0.0755).ToString() + " AZN";
            else if (txtsigortanovu.Text == "Minik avtomobilləri" && txtsigortamuddeti.Text == "60") lbsigorta.Text = "Siğorta: " + (Convert.ToDouble(txtsigortameblegi.Text) * 0.0815).ToString() + " AZN";
            else if (txtsigortanovu.Text == "Minik avtomobilləri" && txtsigortamuddeti.Text == "0") lbsigorta.Text = "Siğorta: 0 AZN";

            else if (txtsigortanovu.Text == "Yük avtomobilləri (3.5 t) və Avtobuslar (8 nəfər)" && txtsigortamuddeti.Text == "12") lbsigorta.Text = "Siğorta: " + (Convert.ToDouble(txtsigortameblegi.Text) * 0.0315).ToString() + " AZN";
            else if (txtsigortanovu.Text == "Yük avtomobilləri (3.5 t) və Avtobuslar (8 nəfər)" && txtsigortamuddeti.Text == "18") lbsigorta.Text = "Siğorta: " + (Convert.ToDouble(txtsigortameblegi.Text) * 0.0365).ToString() + " AZN";
            else if (txtsigortanovu.Text == "Yük avtomobilləri (3.5 t) və Avtobuslar (8 nəfər)" && txtsigortamuddeti.Text == "24") lbsigorta.Text = "Siğorta: " + (Convert.ToDouble(txtsigortameblegi.Text) * 0.0475).ToString() + " AZN";
            else if (txtsigortanovu.Text == "Yük avtomobilləri (3.5 t) və Avtobuslar (8 nəfər)" && txtsigortamuddeti.Text == "30") lbsigorta.Text = "Siğorta: " + (Convert.ToDouble(txtsigortameblegi.Text) * 0.0555).ToString() + " AZN";
            else if (txtsigortanovu.Text == "Yük avtomobilləri (3.5 t) və Avtobuslar (8 nəfər)" && txtsigortamuddeti.Text == "36") lbsigorta.Text = "Siğorta: " + (Convert.ToDouble(txtsigortameblegi.Text) * 0.0610).ToString() + " AZN";
            else if (txtsigortanovu.Text == "Yük avtomobilləri (3.5 t) və Avtobuslar (8 nəfər)" && txtsigortamuddeti.Text == "42") lbsigorta.Text = "Siğorta: " + (Convert.ToDouble(txtsigortameblegi.Text) * 0.0685).ToString() + " AZN";
            else if (txtsigortanovu.Text == "Yük avtomobilləri (3.5 t) və Avtobuslar (8 nəfər)" && txtsigortamuddeti.Text == "48") lbsigorta.Text = "Siğorta: " + (Convert.ToDouble(txtsigortameblegi.Text) * 0.0715).ToString() + " AZN";
            else if (txtsigortanovu.Text == "Yük avtomobilləri (3.5 t) və Avtobuslar (8 nəfər)" && txtsigortamuddeti.Text == "54") lbsigorta.Text = "Siğorta: " + (Convert.ToDouble(txtsigortameblegi.Text) * 0.0755).ToString() + " AZN";
            else if (txtsigortanovu.Text == "Yük avtomobilləri (3.5 t) və Avtobuslar (8 nəfər)" && txtsigortamuddeti.Text == "60") lbsigorta.Text = "Siğorta: " + (Convert.ToDouble(txtsigortameblegi.Text) * 0.0815).ToString() + " AZN";
            else if (txtsigortanovu.Text == "Yük avtomobilləri (3.5 t) və Avtobuslar (8 nəfər)" && txtsigortamuddeti.Text == "0") lbsigorta.Text = "Siğorta: 0 AZN";

            else if (txtsigortanovu.Text == "Yük avtomobilləri (7 tondan yuxarı)" && txtsigortamuddeti.Text == "12") lbsigorta.Text = "Siğorta: " + (Convert.ToDouble(txtsigortameblegi.Text) * 0.0315).ToString() + " AZN";
            else if (txtsigortanovu.Text == "Yük avtomobilləri (7 tondan yuxarı)" && txtsigortamuddeti.Text == "18") lbsigorta.Text = "Siğorta: " + (Convert.ToDouble(txtsigortameblegi.Text) * 0.0415).ToString() + " AZN";
            else if (txtsigortanovu.Text == "Yük avtomobilləri (7 tondan yuxarı)" && txtsigortamuddeti.Text == "24") lbsigorta.Text = "Siğorta: " + (Convert.ToDouble(txtsigortameblegi.Text) * 0.0515).ToString() + " AZN";
            else if (txtsigortanovu.Text == "Yük avtomobilləri (7 tondan yuxarı)" && txtsigortamuddeti.Text == "30") lbsigorta.Text = "Siğorta: " + (Convert.ToDouble(txtsigortameblegi.Text) * 0.0600).ToString() + " AZN";
            else if (txtsigortanovu.Text == "Yük avtomobilləri (7 tondan yuxarı)" && txtsigortamuddeti.Text == "36") lbsigorta.Text = "Siğorta: " + (Convert.ToDouble(txtsigortameblegi.Text) * 0.0680).ToString() + " AZN";
            else if (txtsigortanovu.Text == "Yük avtomobilləri (7 tondan yuxarı)" && txtsigortamuddeti.Text == "42") lbsigorta.Text = "Siğorta: " + (Convert.ToDouble(txtsigortameblegi.Text) * 0.0740).ToString() + " AZN";
            else if (txtsigortanovu.Text == "Yük avtomobilləri (7 tondan yuxarı)" && txtsigortamuddeti.Text == "48") lbsigorta.Text = "Siğorta: " + (Convert.ToDouble(txtsigortameblegi.Text) * 0.0805).ToString() + " AZN";
            else if (txtsigortanovu.Text == "Yük avtomobilləri (7 tondan yuxarı)" && txtsigortamuddeti.Text == "54") lbsigorta.Text = "Siğorta: " + (Convert.ToDouble(txtsigortameblegi.Text) * 0.0855).ToString() + " AZN";
            else if (txtsigortanovu.Text == "Yük avtomobilləri (7 tondan yuxarı)" && txtsigortamuddeti.Text == "60") lbsigorta.Text = "Siğorta: " + (Convert.ToDouble(txtsigortameblegi.Text) * 0.0905).ToString() + " AZN";
            else if (txtsigortanovu.Text == "Yük avtomobilləri (7 tondan yuxarı)" && txtsigortamuddeti.Text == "0") lbsigorta.Text = "Siğorta: 0 AZN";

            else if (txtsigortanovu.Text == "Daşınmaz əmlak" && txtsigortamuddeti.Text == "12") lbsigorta.Text = "Siğorta: " + (Convert.ToDouble(txtsigortameblegi.Text) * 0.0035).ToString() + " AZN";
            else if (txtsigortanovu.Text == "Daşınmaz əmlak" && txtsigortamuddeti.Text == "18") lbsigorta.Text = "Siğorta: " + (Convert.ToDouble(txtsigortameblegi.Text) * 0.0049).ToString() + " AZN";
            else if (txtsigortanovu.Text == "Daşınmaz əmlak" && txtsigortamuddeti.Text == "24") lbsigorta.Text = "Siğorta: " + (Convert.ToDouble(txtsigortameblegi.Text) * 0.0063).ToString() + " AZN";
            else if (txtsigortanovu.Text == "Daşınmaz əmlak" && txtsigortamuddeti.Text == "30") lbsigorta.Text = "Siğorta: " + (Convert.ToDouble(txtsigortameblegi.Text) * 0.0077).ToString() + " AZN";
            else if (txtsigortanovu.Text == "Daşınmaz əmlak" && txtsigortamuddeti.Text == "36") lbsigorta.Text = "Siğorta: " + (Convert.ToDouble(txtsigortameblegi.Text) * 0.0090).ToString() + " AZN";
            else if (txtsigortanovu.Text == "Daşınmaz əmlak" && txtsigortamuddeti.Text == "42") lbsigorta.Text = "Siğorta: " + (Convert.ToDouble(txtsigortameblegi.Text) * 0.0103).ToString() + " AZN";
            else if (txtsigortanovu.Text == "Daşınmaz əmlak" && txtsigortamuddeti.Text == "48") lbsigorta.Text = "Siğorta: " + (Convert.ToDouble(txtsigortameblegi.Text) * 0.0116).ToString() + " AZN";
            else if (txtsigortanovu.Text == "Daşınmaz əmlak" && txtsigortamuddeti.Text == "54") lbsigorta.Text = "Siğorta: " + (Convert.ToDouble(txtsigortameblegi.Text) * 0.0129).ToString() + " AZN";
            else if (txtsigortanovu.Text == "Daşınmaz əmlak" && txtsigortamuddeti.Text == "60") lbsigorta.Text = "Siğorta: " + (Convert.ToDouble(txtsigortameblegi.Text) * 0.0140).ToString() + " AZN";
            else if (txtsigortanovu.Text == "Daşınmaz əmlak" && txtsigortamuddeti.Text == "0") lbsigorta.Text = "Siğorta: 0 AZN";

            for (k1 = 0; k1 < lbsigorta.Text.Length - 4; k1++)
            {
                if (lbsigorta.Text.Substring(k1, 1) == " ")
                {
                    k2 = lbsigorta.Text.Substring(k1 + 1, lbsigorta.Text.Length - k1 - 5);
                    k1 = lbsigorta.Text.Length;
                }
            }

            lbcemixercler.Text = "Cəmi xərclər: "
                               + (Convert.ToDouble(txtavadanliqdeyer.Text) - Convert.ToDouble(txtmalyelesme.Text) //avans
                               + Convert.ToDouble(txtmalyelesme.Text) * Convert.ToDouble(txtkomissiya.Text) / 100 //komisiyya
                               + Convert.ToDouble(txtavadanliqdeyer.Text) * Convert.ToDouble(txtnagdilasdirma.Text) / 100 //nagdilasma
                               + Convert.ToDouble(txtavadanliqdeyer.Text) * Convert.ToDouble(txtkocurme.Text) / 100 //kocurme
                               + Convert.ToDouble(k2) //sigorta
                               + Convert.ToDouble(txtetibarname.Text) * 3 //etibarname
                               + Convert.ToDouble(txtqeyriresmi.Text)).ToString() + " AZN"; //DYP

            label17.Text = (Convert.ToDouble(txtavadanliqdeyer.Text) - Convert.ToDouble(txtmalyelesme.Text)).ToString() + " AZN";
            label20.Text = (Convert.ToDouble(txtmalyelesme.Text) * Convert.ToDouble(txtkomissiya.Text) / 100).ToString() + " AZN";
            label21.Text = (Convert.ToDouble(txtavadanliqdeyer.Text) * Convert.ToDouble(txtnagdilasdirma.Text) / 100).ToString() + " AZN";
            label22.Text = (Convert.ToDouble(txtavadanliqdeyer.Text) * Convert.ToDouble(txtkocurme.Text) / 100).ToString() + " AZN";
            label23.Text = (Convert.ToDouble(txtetibarname.Text) * 3).ToString() + " AZN";
            label25.Text = (Convert.ToDouble(txtqeyriresmi.Text)).ToString() + " AZN";
            label24.Text = (Convert.ToDouble(k2)).ToString() + " AZN";

            txtelealinan.Text = (Convert.ToDouble(txtavadanliqdeyer.Text) - ( //avans
                               +Convert.ToDouble(txtmalyelesme.Text) * Convert.ToDouble(txtkomissiya.Text) / 100 //komisiyya
                               + Convert.ToDouble(txtavadanliqdeyer.Text) * Convert.ToDouble(txtnagdilasdirma.Text) / 100 //nagdilasma
                               + Convert.ToDouble(txtavadanliqdeyer.Text) * Convert.ToDouble(txtkocurme.Text) / 100 //kocurme
                               + Convert.ToDouble(k2) //sigorta
                               + Convert.ToDouble(txtetibarname.Text) * 3 //etibarname
                               + Convert.ToDouble(txtqeyriresmi.Text))).ToString();

            for (k1 = 0; k1 < lbayliqodenis.Text.Length - 4; k1++)
            {
                if (lbayliqodenis.Text.Substring(k1, 1) == ".")
                {
                    k2 = lbayliqodenis.Text.Substring(0, k1 + 3);
                    k1 = lbayliqodenis.Text.Length;
                }

            }
            lbayliqodenis.Text = "Aylıq ödəniş: " + k2 + " AZN";
        }

        private void çıxışToolStripMenuItem_Click(object sender, EventArgs e)
        {
            base.Close();
        }

        private void helpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Hər hansı bir çətinlik barədə, xahiş olunur lizinq mütəxəssislərinə müraciət edəsiniz..(Tel: 012 497-50-17, Ext: 1702, 1703)..");
        }

        private void printToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "") { textBox1.BackColor = Color.Red; return; }

            PrintXercler();
        }

        private void textBox1_MouseClick(object sender, MouseEventArgs e)
        {
            textBox1.BackColor = Color.Gainsboro;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            textBox1.BackColor = Color.Gainsboro;
        }

        private void əlaqəToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Elaqe elaqe = new Elaqe();
            elaqe.ShowDialog();
        }

        private void LizinqCalculator_Load(object sender, EventArgs e)
        {

        }

    }
}
