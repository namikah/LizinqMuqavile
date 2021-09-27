using Nsoft;
using System;
using System.Windows.Forms;

namespace DollarKurs
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void UsdConvert()
        {
            try
            {
                txtAZN.Text = Math.Round((Convert.ToDouble(txtUSD.Text) * Convert.ToDouble(MyChange.Mezenne("USD"))), 4).ToString();
            }
            catch { txtAZN.Text = "0"; }
        }

        private void EurConvert()
        {
            try
            {
                txtAZN2.Text = Math.Round((Convert.ToDouble(txtEUR.Text) * Convert.ToDouble(MyChange.Mezenne("EUR"))), 4).ToString();
            }
            catch { txtAZN2.Text = "0"; }
        }

        private void RubConvert()
        {
            try
            {
                txtAZN3.Text = Math.Round((Convert.ToDouble(txtRUB.Text) * Convert.ToDouble(MyChange.Mezenne("RUB"))), 4).ToString();
            }
            catch { txtAZN3.Text = "0"; }
        }

        private void TryConvert()
        {
            try
            {
                txtAZN4.Text = Math.Round((Convert.ToDouble(txtTRY.Text) * Convert.ToDouble(MyChange.Mezenne("TRY"))), 4).ToString();
            }
            catch { txtAZN4.Text = "0"; }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            btConvert_Click(sender, e);
        }

        private void txtUSD_TextChanged(object sender, EventArgs e)
        {
            UsdConvert();
        }

        private void btConvert_Click(object sender, EventArgs e)
        {
            UsdConvert();
            EurConvert();
            RubConvert();
            TryConvert();
        }

        private void txtEUR_TextChanged(object sender, EventArgs e)
        {
            EurConvert();
        }

        private void txtRUB_TextChanged(object sender, EventArgs e)
        {
            RubConvert();
        }

        private void txtTRY_TextChanged(object sender, EventArgs e)
        {
            TryConvert();
        }
    }
}
