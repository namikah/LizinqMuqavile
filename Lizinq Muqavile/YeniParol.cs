﻿using Nsoft;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Lizinq_Muqavile
{
    public partial class YeniParol : Form
    {
        public YeniParol()
        {
            InitializeComponent();
        }

        private void BtEnterKredit_Click(object sender, EventArgs e)
        {
            MyData.selectCommand("Security", "SELECT * FROM Parol WHERE UserName='" + Environment.UserName + "'");
            MyData.dtmainParol = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainParol);

            if(MyData.dtmainParol.Rows[0]["Parol"].ToString() == txtHazirkiParol.Text)
            {
                MyData.updateCommand("Security", "UPDATE Parol SET Parol='" + txtYeniParol.Text + "' WHERE UserName='" + Environment.UserName + "'");
                MessageBox.Show("Successfully changed","Changed");
            }
            else
            {
                MessageBox.Show("Hazırki parol səhvdir.","Changed");
            }
        }

        private void YeniParol_Load(object sender, EventArgs e)
        {
            txtUserName.Text = Environment.UserName;
        }
    }
}
