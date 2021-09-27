using Nsoft;
using System;
using System.Data;
using System.Windows.Forms;

namespace Lizinq_Muqavile
{
    public partial class Parol : Form
    {
        public Parol()
        {
            InitializeComponent();
        }

        private void TxtParolGiris_Click(object sender, EventArgs e)
        {
            try
            {
                MyData.selectCommand("Security", "SELECT * FROM Parol WHERE UserName='" + Environment.UserName + "'");
                MyData.dtmainParol = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainParol);

                if (MyData.dtmainParol.Rows[0]["UserName"].ToString() == Environment.UserName && MyData.dtmainParol.Rows[0]["Parol"].ToString() == txtParol.Text)
                {
                    MyCheck.Parolicaze = true;
                    base.Close();
                }
            }
            catch
            {
                MyCheck.Parolicaze = false;
                MessageBox.Show("İcazə yoxdur!");
                base.Close();
            }
        }

        private void TxtParol_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                TxtParolGiris_Click(sender, e);
            }
            
        }

        private void Parol_Load(object sender, EventArgs e)
        {

        }
    }
}
