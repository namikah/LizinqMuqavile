using Nsoft;
using System;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;

namespace Lizinq_Muqavile
{
    public partial class TehvilTeslim : Form
    {

        Excel.Application oXL;
        Excel._Workbook oWB;
        Excel._Worksheet oSheet;

        public TehvilTeslim()
        {
            InitializeComponent();
        }

        public void myrefresh()
        {
            try
            {
                string commandText = "SELECT * FROM TehvilTeslim WHERE 1=1";

                commandText += " and a1 like '%" + txtLayihe.Text + "%'";
                commandText += " and a2 like '%" + txtLizinqAlan.Text + "%'";
                commandText += " and a4 like '%" + txtTehvilVeren.Text + "%'";
                commandText += " and a5 like '%" + txtTehvilAlan.Text + "%'";
                commandText += " and a6 like '%" + txtQeyd.Text + "%'";
                commandText += " order by nomre desc";

                MyData.selectCommand("baza.accdb", commandText);
                MyData.dtmain = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmain);
                dataGridView1.DataSource = MyData.dtmain;
            }
            catch { }   
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                MyData.updateCommand("baza.accdb", "UPDATE TehvilTeslim SET "
                                                                                        + "a1 ='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["a1"].Value.ToString() + "',"
                                                                                        + "a2 ='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["a2"].Value.ToString() + "',"
                                                                                        + "a3 ='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["a3"].Value.ToString() + "',"
                                                                                        + "a4 ='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["a4"].Value.ToString() + "',"
                                                                                        + "a5 ='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["a5"].Value.ToString() + "',"
                                                                                        + "a6 ='" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["a6"].Value.ToString() + "'"
                                                                                        + " WHERE nomre Like '" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["nomre"].Value.ToString() + "'");
            }
            catch { MessageBox.Show("Əməliyyat baş tutmadı."); }
        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!MyCheck.davamYesNo()) return;
            try
            {
                MyData.updateCommand("baza.accdb", "DELETE FROM TehvilTeslim WHERE nomre Like '" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["nomre"].Value.ToString() + "'");
            }
            catch { MessageBox.Show("Əməliyyat baş tutmadı."); }
        }

        private void təhvilTəslimYazToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Tehvil_Teslim TehvilTeslimYaz = new Tehvil_Teslim();
            TehvilTeslimYaz.ShowDialog();
            myrefresh();
        }

        private void helpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Elaqe elaqe = new Elaqe();
            elaqe.Show();
        }

        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                File.Copy("Bos.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\TehvilTeslim.xlsx", true);
            }
            catch { MessageBox.Show("Bos.xlsx tapılmadı."); }

            int a, b;

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\TehvilTeslim.xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;

            oSheet.Cells[1, 1] = "№";
            oSheet.Cells[1, 2] = "Layihe";
            oSheet.Cells[1, 3] = "Lizinq Alan";
            oSheet.Cells[1, 4] = "Tarix";
            oSheet.Cells[1, 5] = "Tehvil Veren";
            oSheet.Cells[1, 6] = "Tehvil Alan";
            oSheet.Cells[1, 7] = "Qeydlər";

            for (a = 0; a < dataGridView1.Rows.Count; a++)
            {
                for (b = 0; b < 7; b++)
                {
                    oSheet.Cells[a + 2, b + 1] = dataGridView1.Rows[a].Cells[b].Value.ToString();

                }
                oSheet.Range["A" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["B" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["C" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["D" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["E" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["F" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["G" + (a + 2)].Borders.LineStyle = Excel.Constants.xlSolid;

            }


            oSheet.Range["A" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["B" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["C" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["D" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["E" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["F" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["G" + 1].Borders.LineStyle = Excel.Constants.xlSolid;

            oSheet.Columns.AutoFit();
            oSheet.Rows.AutoFit();
        }

        private void Button15_Click(object sender, EventArgs e)
        {
            myrefresh();
        }

        private void Btrefresh_Click(object sender, EventArgs e)
        {
            myrefresh();
        }

        private void TehvilTeslim_Load(object sender, EventArgs e)
        {
            myrefresh();
        }

        private void TxtLayihe_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                myrefresh();
            }
        }

        private void TxtLizinqAlan_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                myrefresh();
            }
        }

        private void TxtTehvilVeren_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                myrefresh();
            }
        }

        private void TxtTehvilAlan_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                myrefresh();
            }
        }

        private void TxtQeyd_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                myrefresh();
            }
        }
    }
}
