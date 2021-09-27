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
    public partial class Qeydler : Form
    {
        public Qeydler()
        {
            InitializeComponent();
        }

        private void myrefresh()
        {
            MyData.selectCommand("baza.accdb", "Select * from Qeydler");
            MyData.dtmain = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmain);
            dataGridView1.DataSource = MyData.dtmain;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                string commandText = "SELECT * FROM Qeydler WHERE  1=1";
                commandText += " and c1 like " + "'%" + textBox1.Text + "%'";
                commandText += " or c2 like " + "'%" + textBox1.Text + "%'";

                MyData.selectCommand("baza.accdb", commandText);
                MyData.dtmain = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmain);
                dataGridView1.DataSource = MyData.dtmain;

            }
            catch { };
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    string commandText = "SELECT * FROM Qeydler WHERE  1=1";
                    commandText += " and c1 like " + "'%" + textBox1.Text + "%'";
                    commandText += " or c2 like " + "'%" + textBox1.Text + "%'";

                    MyData.selectCommand("baza.accdb", commandText);
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
                MyData.updateCommand("baza.accdb", "UPDATE Qeydler SET "
                                                                                     + "c1 ='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString() + "',"
                                                                                     + "c2 ='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[2].Value.ToString() + "'"
                                                                                     + " WHERE Код Like'" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString() + "'");
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
                MyData.deleteCommand("baza.accdb", "DELETE FROM Qeydler WHERE Код Like '" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString() + "'");
                myrefresh();
            }
            catch { MessageBox.Show("Əməliyyat baş tutmadı."); }
        }

        private static CultureInfo ci = new CultureInfo("AZ");

        private void button1_Click(object sender, EventArgs e)
        {
            try { richTextBox1.Text = richTextBox1.Text.ToUpper(ci); }
            catch { }

            try
            {
                MyData.insertCommand("baza.accdb", "insert into Qeydler (c1,c2)values('"

                                                                                                    + dttarix.Text + "','"
                                                                                                    + richTextBox1.Text + "')");
                myrefresh();
            }
            catch { }
        }

        private void Qeydler_Load(object sender, EventArgs e)
        {
            myrefresh();
        }

        private void richTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            //SendKeys.Send("{CAPSLOCK}");
           // e.KeyChar = char.ToUpper(e.KeyChar);
           
        }

    }
}
