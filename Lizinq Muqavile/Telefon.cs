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
using Nsoft;

namespace Lizinq_Muqavile
{
    public partial class Telefon : Form
    {
        public Telefon()
        {
            InitializeComponent();
        }

        private void myrefresh()
        {
            MyData.selectCommand("baza.accdb", "Select * from Telefon");
            MyData.dtmain = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmain);
            dataGridView1.DataSource = MyData.dtmain;
        }

        private void Telefon_Load(object sender, EventArgs e)
        {
            MyChange.SetKeyboardLayout(MyChange.GetInputLanguageByName("AZ"));
            myrefresh();
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    MyData.selectCommand("baza.accdb", "SELECT * FROM Telefon WHERE c1 like '%" + textBox1.Text + "%'");
                    MyData.dtmain = new DataTable();
                    MyData.oledbadapter1.Fill(MyData.dtmain);
                    dataGridView1.DataSource = MyData.dtmain;
                }
                catch { };

            }
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView1.EditMode = DataGridViewEditMode.EditProgrammatically;

            try
            {
                MyData.updateCommand("baza.accdb", "UPDATE Telefon SET "
                                                                                     + "c2 ='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c2"].Value.ToString() + "',"
                                                                                     + "c3 ='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c3"].Value.ToString() + "'"
                                                                                     + " WHERE c1 Like '%" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c1"].Value.ToString() + "%'");
            }
            catch { MessageBox.Show("Əməliyyat baş tutmadı."); }
        }

        private void editToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.EditMode = DataGridViewEditMode.EditOnEnter;
        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                MyData.deleteCommand("baza.accdb", "DELETE FROM Telefon WHERE c1 Like '%" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["c1"].Value.ToString() + "%'");
                myrefresh();
            }
            catch { MessageBox.Show("Əməliyyat baş tutmadı."); }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (MyData.dtmain.Rows.Count == dataGridView1.Rows.Count - 1) return;

            MyData.deleteCommand("baza.accdb", "insert into Telefon (c1,c2,c3)values("

                                                                                                 + "'" + dataGridView1.Rows[dataGridView1.Rows.Count - 2].Cells["c1"].Value.ToString() + "',"
                                                                                                + "'" + dataGridView1.Rows[dataGridView1.Rows.Count - 2].Cells["c2"].Value.ToString() + "',"
                                                                                                + "'" + dataGridView1.Rows[dataGridView1.Rows.Count - 2].Cells["c3"].Value.ToString() + "')");
            MessageBox.Show("Əlavə edildi.");
            myrefresh();
        }


    }
}
