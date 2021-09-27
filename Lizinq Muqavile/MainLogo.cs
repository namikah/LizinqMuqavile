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
    public partial class MainLogo : Form
    {
        public MainLogo()
        {
            InitializeComponent();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            base.Opacity = base.Opacity + 0.04;
            pictureBox1.Top += 4;
            label1.Left += 12;
            label2.Left -= 12;
            //if (pictureBox1.Top > base.Height / 7) { timer1.Enabled = false; timer2.Enabled = true; }
            if (label2.Left < 90) { timer1.Enabled = false; timer2.Enabled = true; }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            base.Opacity = base.Opacity - 0.02;
            if (base.Opacity < 0.02) { timer2.Enabled = false; base.Close(); }
        }

        private void MainLogo_MouseClick(object sender, MouseEventArgs e)
        {
            base.Close();
        }

        private void label1_Click(object sender, EventArgs e)
        {
            base.Close();
        }

        private void label2_Click(object sender, EventArgs e)
        {
            base.Close();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            base.Close();
        }

        private void MainLogo_Load(object sender, EventArgs e)
        {

        }
    }
}
