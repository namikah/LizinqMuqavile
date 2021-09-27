using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;

namespace Lizinq_Muqavile
{
    public partial class Rekvizitler : Form
    {
        public Rekvizitler()
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

        private void FindAndReplace(Word.Application word, object findText, object replaceText)
        {
            word.Selection.Find.ClearFormatting();
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = true;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 1;
            object wrap = 2;

            word.Selection.Find.Execute(ref findText, ref matchCase,
            ref matchWholeWord, ref matchWildCards, ref matchSoundsLike,
            ref matchAllWordForms, ref forward, ref wrap, ref format,
            ref replaceText, ref replace, ref matchKashida,
            ref matchDiacritics,
            ref matchAlefHamza, ref matchControl);
        }

        public void WordDoc()
        {
            try { File.Copy("New Emphty.doc", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Rekvizitlər.doc", true); }
            catch { MessageBox.Show("'\\192.168.10.5\\Common\\AGLizinq\\New Emphty.doc' tapılmadı."); }

            object FileName = "C:\\Users\\" + Environment.UserName + "\\Desktop\\Rekvizitlər.doc";
            Microsoft.Office.Interop.Word._Application oWord;
            object oMissing = Type.Missing;
            oWord = new Microsoft.Office.Interop.Word.Application();
            oWord.Visible = true;
            oWord.Documents.Open(FileName);
            oWord.Selection.TypeText("“AGLİZİNQ” QAPALI SƏHMDAR CƏMİYYƏTİNİN" + Environment.NewLine + "BANK REKVİZİTLƏRİ" + Environment.NewLine + Environment.NewLine + Environment.NewLine);
            oWord.Selection.TypeText(richTextBox1.Text + Environment.NewLine + Environment.NewLine + richTextBox2.Text);
            //oWord.PrintOut();
            oWord.ActiveDocument.Save();
            //oWord.Quit();
        }

        private void printToolStripMenuItem_Click(object sender, EventArgs e)
        {
            WordDoc();
        }

        private void çıxışToolStripMenuItem_Click(object sender, EventArgs e)
        {
            base.Close();
        }

        private void infoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Elaqe elaqe = new Elaqe();
            elaqe.ShowDialog();
        }
    }
}
