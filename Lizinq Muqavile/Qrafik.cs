using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Lizinq_Muqavile
{
    public partial class Qrafik : Form
    {
        public Qrafik()
        {
            InitializeComponent();
        }

        private void Qrafik_Load(object sender, EventArgs e)
        {

        }

        private void txtmebleg_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) txtmebleg.Text = Convert.ToInt32(Convert.ToDouble(txtodenis.Text) * (1 - 1 / Math.Pow((1 + Convert.ToDouble(txtfaiz.Text) / 100 / 12), Convert.ToDouble(txtmuddet.Text))) / (Convert.ToDouble(txtfaiz.Text) / 100 / 12)).ToString();

        }

        private void txtodenis_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) txtodenis.Text = Convert.ToInt32(Convert.ToDouble(txtmebleg.Text) * (Convert.ToDouble(txtfaiz.Text) / 100 / 12) / (1 - 1 / Math.Pow((1 + Convert.ToDouble(txtfaiz.Text) / 100 / 12), Convert.ToDouble(txtmuddet.Text)))).ToString();

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            base.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }
    }
}
