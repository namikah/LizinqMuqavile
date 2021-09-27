using System;
using System.Data;
using System.Windows.Forms;
using Nsoft;

namespace Lizinq_Muqavile
{
    public partial class Lisenziya : Form
    {
        public Lisenziya()
        {
            InitializeComponent();
        }

        private void LisenziyaRefresh()  //---------------rasxodun nomresinin load olunmasi-------------------------------
        {
            try
            {
                MyData.selectCommand("baza.accdb", "Select * From Lisenziya");
                MyData.dtmain = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmain);

                DateTime dt = DateTime.Now;
                DateTime dt2 = Convert.ToDateTime(MyData.dtmain.Rows[0]["a2"]);

                label1.Text = "Lisenziya vaxtı: " + dt2.ToShortDateString() + " (" + (dt2-dt).Days + " gün qalıb)";
            }
            catch { label1.Text = "Lisenziya Yoxdur!"; }
        }

        private void Lisenziya_Load(object sender, EventArgs e)
        {
            LisenziyaRefresh();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try { if (textBox1.Text.Substring(0, 8) != "Nhl99nhl") { MessageBox.Show("Kod Yalnışdır"); return; } }
            catch { MessageBox.Show("Kod Yalnışdır"); return; }

            DateTime dt = DateTime.Now;

            MyData.updateCommand("baza.accdb", "UPDATE Lisenziya SET "
                                                                                 + "a1 ='" + dt.ToShortDateString() + "',"
                                                                                 + "a2 ='" + textBox1.Text.Substring(textBox1.Text.Length - 10, 10) + "'");
            LisenziyaRefresh();

            MessageBox.Show("Müvəffəqiyyətlə yeniləndi.");
            base.Close();
        }

    }
}
