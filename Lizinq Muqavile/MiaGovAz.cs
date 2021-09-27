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
    public partial class MiaGovAz : Form
    {
        public MiaGovAz()
        {
            InitializeComponent();
        }

        private void MiaGovAz_Load(object sender, EventArgs e)
        {
           webBrowser1.Navigate("http://mia.gov.az/?/az/driverlicense/");
        }
    }
}
