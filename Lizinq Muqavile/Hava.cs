using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using Nsoft;

namespace Lizinq_Muqavile
{
    public partial class Hava : Form
    {
        public Hava()
        {
            InitializeComponent();
        }

        private void Hava_Load(object sender, EventArgs e)
        {
            label1.Text = "Bakı, Azərbaycan" + Environment.NewLine + Environment.NewLine + MyChange.HavaBaku() + " °C";
        }

        private void label1_Click(object sender, EventArgs e)
        {
            
        }
    }
}
