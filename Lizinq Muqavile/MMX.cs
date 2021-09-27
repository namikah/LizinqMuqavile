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
using System.Net.Mail;
using System.Net;
using System.IO;
using System.Net.Sockets;
using System.Web;
using Nsoft;

namespace Lizinq_Muqavile
{
    public partial class MMX : Form
    {
        public MMX()
        {
            InitializeComponent();
        }

        public void myrefresh()
        {
            try
            {
                MyData.selectCommand("baza.accdb", "SELECT * FROM MMX WHERE a2 like '%" + textBox1.Text
                    + "%' or a7 like '%" + textBox1.Text
                    + "%' or a1 like '%" + textBox1.Text + "%' order by nomre desc");

                MyData.dtmain = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmain);
                dataGridView1.DataSource = MyData.dtmain;

                double k = 0;
                for (int i = 0; i < MyData.dtmain.Rows.Count; i++) k = Math.Round(Convert.ToDouble(k) + Convert.ToDouble(MyData.dtmain.Rows[i]["a10"]),2);
                btcemi.Text = "Cəmi: " + k.ToString() + " AZN";
            }
            catch { }

        }

        private void MMX_Load(object sender, EventArgs e)
        {
            myrefresh();
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                myrefresh();
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            myrefresh();
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                MyData.selectCommand("baza.accdb", "SELECT c4 FROM etibarnameneqliyyat WHERE c1 like '%" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["a2"].Value + "%'");
                MyData.dtmain = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmain);

                listBox3.DataSource = MyData.dtmain;
                listBox3.DisplayMember = "c4";
                //listBox3.ValueMember = "c14";

                MyData.selectCommand("baza.accdb","SELECT a1 FROM etibarnamesurucu WHERE a5 like '%" + MyData.dtmain.Rows[0]["c4"] + "%'");
                MyData.dtmainSuruculer = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainSuruculer);

                listBox1.DataSource = MyData.dtmainSuruculer;
                listBox1.DisplayMember = "a1";
                //listBox1.ValueMember = "Kod";

                if (listBox1.Items.Count > 0) listBox1.SetSelected(0, true);
            }
            catch { }

            try { seçilmişProtokoluTəhvilVerToolStripMenuItem.Text = "(" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["a2"].Value + ") Seçilmiş protokolu Təhvil ver"; }
            catch { }

            try { hamısınıTəhvilVerToolStripMenuItem.Text = "(" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["a2"].Value + ") Bütün protokolları təhvil ver"; }
            catch { }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                base.Text = listBox1.Items[listBox1.SelectedIndex].ToString();
                MyData.selectCommand("baza.accdb", "SELECT * FROM Telefon WHERE c1 like '%" + ((DataRowView)listBox1.SelectedItem).Row[0] + "%'");
                MyData.dtmain = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmain);

                listBox2.Items.Clear();
                listBox2.Items.Add(MyData.dtmain.Rows[0]["c2"].ToString());
                listBox2.Items.Add(MyData.dtmain.Rows[0]["c3"].ToString());
            }
            catch { }
        
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if(textBox1.Text.Length > 4) myrefresh();   
        }

        private void telefonKitabçasıToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Telefon telefon = new Telefon();
            telefon.Show();
        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                MyData.deleteCommand("baza", "DELETE FROM MMX WHERE nomre Like '" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["nomre"].Value.ToString() + "'");

                MessageBox.Show("Əməliyyat yerinə yetirildi.");
                myrefresh();
            }
            catch { MessageBox.Show("Əməliyyat baş tutmadı."); }
        }

        private void editToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.EditMode = DataGridViewEditMode.EditOnEnter;
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView1.EditMode = DataGridViewEditMode.EditProgrammatically;
            try
            {
                MyData.updateCommand("baza.accdb", "UPDATE MMX SET "
                                                                                     + "a1 ='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["a1"].Value.ToString() + "',"
                                                                                     + "a2 ='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["a2"].Value.ToString() + "',"
                                                                                     + "a3 ='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["a3"].Value.ToString() + "',"
                                                                                     + "a4 ='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["a4"].Value.ToString() + "',"
                                                                                     + "a5 ='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["a5"].Value.ToString() + "',"
                                                                                     + "a6 ='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["a6"].Value.ToString() + "',"
                                                                                     + "a7 ='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["a7"].Value.ToString() + "',"
                                                                                     + "a8 ='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["a8"].Value.ToString() + "',"
                                                                                     + "a9 ='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["a9"].Value.ToString() + "',"
                                                                                     + "a10 ='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["a10"].Value.ToString() + "'"
                                                                                     + " WHERE nomre Like '" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["nomre"].Value.ToString() + "'");
            }
            catch { MessageBox.Show("Əməliyyat baş tutmadı."); }
        }

        private void hamısınıTəhvilVerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DateTime dt = DateTime.Now;
            try
            {
                MyData.updateCommand("baza.accdb", "UPDATE MMX SET "
                                                                                     + "a8 ='Bəli',"
                                                                                     + "a9 ='Bəli " + dt.Date + "'"
                                                                                     + " WHERE NOT a9 Like '%Bəli%' and a2 Like '%" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["a2"].Value.ToString() + "%'");

                myrefresh();
            }
            catch { MessageBox.Show("Əməliyyat baş tutmadı."); }
        }

        private void seçilmişProtokoluTəhvilVerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DateTime dt = DateTime.Now;
            try
            {
                MyData.updateCommand("baza.accdb", "UPDATE MMX SET "
                                                                                     + "a8 ='Bəli',"
                                                                                     + "a9 ='Bəli " + dt.Date + "'"
                                                                                     + " WHERE NOT a9 Like '%Bəli%' and nomre Like '%" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["nomre"].Value.ToString() + "%'");

                myrefresh();
            }
            catch { MessageBox.Show("Əməliyyat baş tutmadı."); }
        }

        private void sürücüyəBildirişYazToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MMXmelumat MMXmelumat = new MMXmelumat();
            MMXmelumat.Show();
        }

        private void mMXToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MMXmelumat MMXmelumat = new MMXmelumat();
            MMXmelumat.Show();
        }

        private void bNAToolStripMenuItem_Click(object sender, EventArgs e)
        {
            BNAadd bnaadd = new BNAadd();
            bnaadd.Show();
        }

    }
}
