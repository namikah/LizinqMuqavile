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
using System.Globalization;
using System.IO;
using System.Net;
using System.Web;

namespace Lizinq_Muqavile
{
    public partial class Sekiller : Form
    {
        public Sekiller()
        {
            InitializeComponent();
        }

        OleDbDataAdapter oledbadapter1; ////neqliyyat vasiteleri
        OleDbConnection oledbconnection1;
        DataTable dtmain;

        private void CreateSqlConnection()
        {
            oledbconnection1 = new OleDbConnection();
            oledbadapter1 = new OleDbDataAdapter();
            dtmain = new DataTable();
            oledbconnection1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='baza.accdb'";
        }    //elaqe yaratmaq

        void RefreshAvtomobiller()
        {
            OleDbConnection cn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='baza.accdb'");
            OleDbCommand cmd = new OleDbCommand();
            OleDbDataAdapter da = new OleDbDataAdapter();
            DataTable dtb = new DataTable();

            dataGridView4.RowTemplate.Height = 200;
            da.SelectCommand = new OleDbCommand();
            da.SelectCommand.Connection = cn;
            da.SelectCommand.CommandText = "Select * from AvtoMelumat";
            dtb.Clear();
            da.Fill(dtb);
            dataGridView4.DataSource = dtb;
            dataGridView4.Columns["Images"].Width = 200;
            for (int i = 0; i < dataGridView4.Columns.Count; i++)
            {
                if (dataGridView4.Columns[i] is DataGridViewImageColumn)
                {
                    ((DataGridViewImageColumn)dataGridView4.Columns[i]).ImageLayout = DataGridViewImageCellLayout.Stretch;
                }
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            int i = 0, s = 0;
            string k = t1.Text;
            OleDbConnection cn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='baza.accdb'");
            OleDbCommand cmd = new OleDbCommand();

            CreateSqlConnection();
            oledbadapter1.SelectCommand = new OleDbCommand();
            oledbadapter1.SelectCommand.Connection = oledbconnection1;
            oledbadapter1.SelectCommand.CommandText = "Select * from etibarnameneqliyyat WHERE c1 Like '%" + t1.Text + "%'";
            dtmain.Clear();
            oledbadapter1.Fill(dtmain);

            try
            {
                t1.Text = dtmain.Rows[0]["c1"].ToString();
                k = "Markası - " + dtmain.Rows[s]["c2"].ToString() + " (" + "Nömrə nişanı - " + dtmain.Rows[0]["c1"].ToString() + ", Buraxılış ili - " + dtmain.Rows[0]["c6"].ToString();
            }
            catch { }

            cn.Open();
            cmd.Connection = cn;
            cmd.CommandText = "INSERT INTO AvtoMelumat(c1,c2) Values('" + k.ToString() + "',@Images)";
            MemoryStream stream = new MemoryStream();
            pb1.Image.Save(stream, System.Drawing.Imaging.ImageFormat.Jpeg);
            byte[] pic = stream.ToArray();
            cmd.Parameters.AddWithValue("@Images", pic);
            i = cmd.ExecuteNonQuery();
            cn.Close();
            if (i > 0)
            {
                MessageBox.Show("Insert Success " + i);
            }

            RefreshAvtomobiller();
        }

        private void button16_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "jpg|*.jpg";
            DialogResult res = openFileDialog1.ShowDialog();
            if (res == DialogResult.OK)
            {
                pb1.Image = Image.FromFile(openFileDialog1.FileName);
            }
        }

        private void Sekiller_Load(object sender, EventArgs e)
        {
            RefreshAvtomobiller();
        }
    }
}
