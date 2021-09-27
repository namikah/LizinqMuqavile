using Lizinq_Muqavile;
using System;
using System.Data;
using System.Windows.Forms;

namespace Nsoft
{
    class MyCheck
    {
        public static bool Parolicaze = false;

        public static Boolean davamYesNo(string MessageText)
        {
            try
            {
                DialogResult result2 = MessageBox.Show(MessageText, "info", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                if (result2 == DialogResult.Yes) return true;
                else return false;
            }
            catch { return false; }
        }

        public static Boolean davamYesNo()
        {
            try
            {
                DialogResult result2 = MessageBox.Show("Davam etmək istəyirsiniz?", "info", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                if (result2 == DialogResult.Yes) return true;
                else return false;
            }
            catch { return false; }
        }

        public static Boolean ParolYesNo()
        {
            try
            {
                MyData.selectCommand("Security", "SELECT * FROM Parol WHERE UserName='" + Environment.UserName + "'");
                MyData.dtmainParol = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainParol);

                if (MyData.dtmainParol.Rows[0]["UserName"].ToString() == Environment.UserName &&
                    MyData.dtmainParol.Rows[0]["Parol"].ToString() == "") return true;
                else return false;
            }
            catch { return false; }
        }

        public static Boolean ParolAdminYesNo()
        {
            try
            {
                MyData.selectCommand("Security", "SELECT * FROM Parol WHERE UserName='" + Environment.UserName + "'");
                MyData.dtmainParol = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainParol);

                if (MyData.dtmainParol.Rows[0]["UserName"].ToString() == Environment.UserName &&
                    MyData.dtmainParol.Rows[0]["Status"].ToString() == "Admin") return true;
                else { MessageBox.Show("İcazə yoxdur!"); return false; }
            }
            catch { return false; }
        }

        public static void ParolYoxla()
        {
            if (ParolYesNo())
            {
                Parolicaze = true;
                return;
            }

            if (!Parolicaze)
            {
                Parol parol = new Parol();
                parol.ShowDialog();
            }
        }


        public static Boolean LisenziyaYoxla()
        {
            try
            {
                MyData.selectCommand("baza.accdb", "Select * From Lisenziya");
                MyData.dtmainLisenziya = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainLisenziya);

                DateTime dt = DateTime.Now;
                DateTime dt2 = Convert.ToDateTime(MyData.dtmainLisenziya.Rows[0]["a2"]);

                if (dt <= dt2) return true;
                else return false;
            }
            catch { return false; }
        }
    }
}
