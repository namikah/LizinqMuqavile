using Nsoft;
using System;
using System.Data;
using System.Windows.Forms;

namespace Lizinq_Muqavile
{
    public partial class Emeliyyatlar : Form
    {
        public Emeliyyatlar()
        {
            InitializeComponent();
        }

        private void myrefresh()  //---------------rasxodun nomresinin load olunmasi--------------
        {
            try
            {
                string commandText = "SELECT * FROM Emeliyyatlar WHERE 1=1";
                commandText += " and a2 Like '%" + txtAxtar.Text + "%'";
                commandText += " and a1 between #" + dtBaslama.Value.ToString("yyyy-MM-dd") + "# and #" + dtBitme.Value.AddDays(1).ToString("yyyy-MM-dd") + "# order by Kod desc";
                MyData.selectCommand("baza.accdb", commandText);
                MyData.dtmain = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmain);
                dataGridView1.DataSource = MyData.dtmain;
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void Elaveler_Load(object sender, EventArgs e)
        {
            myrefresh();
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView1.EditMode = DataGridViewEditMode.EditProgrammatically;

            try
            {
                MyData.updateCommand("baza.accdb", "UPDATE Emeliyyatlar SET a1 ='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Tarix"].Value.ToString() + "',"
                    + "a2 ='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Info"].Value.ToString() + "'"
                    + " WHERE Kod Like'" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["Kod"].Value.ToString() + "'");
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void editToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.EditMode = DataGridViewEditMode.EditOnEnter;
        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult result2 = MessageBox.Show("Davam etmək istəyirsiniz?", "?", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (result2 == DialogResult.No) { return; }

            try
            {
                MyData.deleteCommand("baza.accdb", "DELETE FROM Emeliyyatlar WHERE Kod Like '" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString() + "'");
                myrefresh();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message);}

            myrefresh();
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                computerToolStripMenuItem.Text = "ComputerName (" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[3].Value.ToString() + ")";
            }
            catch { }
        }

        private void DtBitme_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                myrefresh();
            }
        }

        private void DtBaslama_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                myrefresh();
            }
        }

        private void TxtAxtar_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
            {
                myrefresh();
            }
        }
    }
}
