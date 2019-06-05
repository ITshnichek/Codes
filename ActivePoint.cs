using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows.Forms;
using PackingTool.Properties;

namespace PackingTool
{
    public partial class ActivePoint : Form
    {
        int y = 0;
        private string yan3;
        Button cb = null;
        Control c;
        int l = 1;
        MySqlDataReader reader;
        MySqlDataReader reader1;

        public ActivePoint()
        {
            InitializeComponent();
        }

        string odi;

        private void ActivePoint_Load(object sender, EventArgs e)
        {
            try
            {
                textBox1.Focus();
                label5.Text = Settings.Default.referencia;

                m2:
                foreach (Control c in this.Controls)
                {
                    jio();
                    cb = c as Button;
                    string hoho = Convert.ToString(c);
                    string hoh = "System.Windows.Forms.Button, Text: " + Convert.ToString(l) + "";
                    if (hoho == hoh)
                    {
                        if (l != 51 && odi == "1")
                        {
                            reader.Close();
                            l++;
                            c.BackColor = Color.Blue;
                        }
                        else
                        {
                            l++;
                        }

                    }
                    if (l == 51)
                    {
                        reader.Close();
                        if (button1.BackColor != Color.Blue &&
                            button2.BackColor != Color.Blue &&
                            button3.BackColor != Color.Blue &&
                            button4.BackColor != Color.Blue &&
                            button5.BackColor != Color.Blue &&
                            button6.BackColor != Color.Blue &&
                            button7.BackColor != Color.Blue &&
                            button8.BackColor != Color.Blue &&
                            button9.BackColor != Color.Blue &&
                            button10.BackColor != Color.Blue &&
                            button11.BackColor != Color.Blue &&
                            button12.BackColor != Color.Blue &&
                            button13.BackColor != Color.Blue &&
                            button14.BackColor != Color.Blue &&
                            button15.BackColor != Color.Blue &&
                            button16.BackColor != Color.Blue &&
                            button17.BackColor != Color.Blue &&
                            button18.BackColor != Color.Blue &&
                            button19.BackColor != Color.Blue &&
                            button20.BackColor != Color.Blue &&
                            button21.BackColor != Color.Blue &&
                            button22.BackColor != Color.Blue &&
                            button23.BackColor != Color.Blue &&
                            button24.BackColor != Color.Blue &&
                            button25.BackColor != Color.Blue &&
                            button26.BackColor != Color.Blue &&
                            button27.BackColor != Color.Blue &&
                            button28.BackColor != Color.Blue &&
                            button29.BackColor != Color.Blue &&
                            button30.BackColor != Color.Blue &&
                            button31.BackColor != Color.Blue &&
                            button32.BackColor != Color.Blue &&
                            button33.BackColor != Color.Blue &&
                            button34.BackColor != Color.Blue &&
                            button35.BackColor != Color.Blue &&
                            button36.BackColor != Color.Blue &&
                            button37.BackColor != Color.Blue &&
                            button38.BackColor != Color.Blue &&
                            button39.BackColor != Color.Blue &&
                            button40.BackColor != Color.Blue &&
                            button41.BackColor != Color.Blue &&
                            button42.BackColor != Color.Blue &&
                            button43.BackColor != Color.Blue &&
                            button44.BackColor != Color.Blue &&
                            button45.BackColor != Color.Blue &&
                            button46.BackColor != Color.Blue &&
                            button47.BackColor != Color.Blue &&
                            button48.BackColor != Color.Blue &&
                            button49.BackColor != Color.Blue &&
                            button50.BackColor != Color.Blue)

                        {
                            referencia dbs = new referencia();
                            dbs.Show();
                            this.Close();
                        }
                        return;
                    }
                }
                if (l <= 51)
                {
                    goto m2;
                }
                if (y == 1)
                {
                    referencia dbs = new referencia();
                    dbs.Show();
                    this.Close();
                    y = 0;
                }
                else
                {
                    y = 0;
                }
            }
            catch (Exception)
            {
                referencia dbs = new referencia();
                dbs.Show();
                this.Close();
                y = 0;
            }


        }

        public void jio()
        {

            try
            {
                string connectionString = @"" + Settings.Default.packingConnectionString + "";
                using (MySqlConnection conn = new MySqlConnection(connectionString))
                {
                    conn.Open();
                    using (
                        MySqlCommand cmd =
                            new MySqlCommand(
                                "SELECT p" + l + " FROM " + Settings.Default.project + "_" +
                                Settings.Default.semeistvo + " WHERE referencia='" +
                                Settings.Default.referencia + "' ;", conn))
                    {
                        reader = cmd.ExecuteReader();
                        int i = 0;
                        try
                        {
                            while (reader.Read())
                            {
                                odi = reader[0].ToString();
                            }
                        }
                        catch (Exception)
                        {

                        }
                    }
                }
            }
            catch (Exception)
            {

                y = 1;
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {

            if (e.KeyCode == Keys.Enter)

            {
                if (textBox1.Text == "ACTIVEPOINT1" && button1.BackColor != Color.DimGray)
                {
                    button1.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT2" && button2.BackColor != Color.DimGray)
                {
                    button2.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT3" && button3.BackColor != Color.DimGray)
                {
                    button3.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT4" && button4.BackColor != Color.DimGray)
                {
                    button4.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT5" && button5.BackColor != Color.DimGray)
                {
                    button5.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT6" && button6.BackColor != Color.DimGray)
                {
                    button6.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT7" && button7.BackColor != Color.DimGray)
                {
                    button7.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT8" && button8.BackColor != Color.DimGray)
                {
                    button8.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT9" && button9.BackColor != Color.DimGray)
                {
                    button9.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT10" && button10.BackColor != Color.DimGray)
                {
                    button10.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT11" && button11.BackColor != Color.DimGray)
                {
                    button11.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT12" && button12.BackColor != Color.DimGray)
                {
                    button12.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT13" && button13.BackColor != Color.DimGray)
                {
                    button13.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT14" && button14.BackColor != Color.DimGray)
                {
                    button14.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT15" && button15.BackColor != Color.DimGray)
                {
                    button15.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT16" && button16.BackColor != Color.DimGray)
                {
                    button16.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT17" && button17.BackColor != Color.DimGray)
                {
                    button17.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT18" && button18.BackColor != Color.DimGray)
                {
                    button18.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT19" && button19.BackColor != Color.DimGray)
                {
                    button19.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT20" && button20.BackColor != Color.DimGray)
                {
                    button20.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT21" && button21.BackColor != Color.DimGray)
                {
                    button21.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT22" && button22.BackColor != Color.DimGray)
                {
                    button22.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT23" && button23.BackColor != Color.DimGray)
                {
                    button23.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT24" && button24.BackColor != Color.DimGray)
                {
                    button24.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT25" && button25.BackColor != Color.DimGray)
                {
                    button25.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT26" && button26.BackColor != Color.DimGray)
                {
                    button26.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT27" && button27.BackColor != Color.DimGray)
                {
                    button27.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT28" && button28.BackColor != Color.DimGray)
                {
                    button28.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT29" && button29.BackColor != Color.DimGray)
                {
                    button29.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT30" && button30.BackColor != Color.DimGray)
                {
                    button30.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT31" && button31.BackColor != Color.DimGray)
                {
                    button31.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT32" && button32.BackColor != Color.DimGray)
                {
                    button32.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT33" && button33.BackColor != Color.DimGray)
                {
                    button33.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT34" && button34.BackColor != Color.DimGray)
                {
                    button34.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT35" && button35.BackColor != Color.DimGray)
                {
                    button35.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT36" && button36.BackColor != Color.DimGray)
                {
                    button36.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT37" && button37.BackColor != Color.DimGray)
                {
                    button37.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT38" && button38.BackColor != Color.DimGray)
                {
                    button38.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT39" && button39.BackColor != Color.DimGray)
                {
                    button39.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT40" && button40.BackColor != Color.DimGray)
                {
                    button40.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT41" && button41.BackColor != Color.DimGray)
                {
                    button41.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT42" && button42.BackColor != Color.DimGray)
                {
                    button42.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT43" && button43.BackColor != Color.DimGray)
                {
                    button43.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT44" && button44.BackColor != Color.DimGray)
                {
                    button44.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT45" && button45.BackColor != Color.DimGray)
                {
                    button45.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT46" && button46.BackColor != Color.DimGray)
                {
                    button46.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT47" && button47.BackColor != Color.DimGray)
                {
                    button47.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT48" && button48.BackColor != Color.DimGray)
                {
                    button48.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT49" && button49.BackColor != Color.DimGray)
                {
                    button49.BackColor = Color.Red;
                }
                else if (textBox1.Text == "ACTIVEPOINT50" && button4.BackColor != Color.DimGray)
                {
                    button50.BackColor = Color.Red;
                }
                if (button1.BackColor != Color.Blue &&
                    button2.BackColor != Color.Blue &&
                    button3.BackColor != Color.Blue &&
                    button4.BackColor != Color.Blue &&
                    button5.BackColor != Color.Blue &&
                    button6.BackColor != Color.Blue &&
                    button7.BackColor != Color.Blue &&
                    button8.BackColor != Color.Blue &&
                    button9.BackColor != Color.Blue &&
                    button10.BackColor != Color.Blue &&
                    button11.BackColor != Color.Blue &&
                    button12.BackColor != Color.Blue &&
                    button13.BackColor != Color.Blue &&
                    button14.BackColor != Color.Blue &&
                    button15.BackColor != Color.Blue &&
                    button16.BackColor != Color.Blue &&
                    button17.BackColor != Color.Blue &&
                    button18.BackColor != Color.Blue &&
                    button19.BackColor != Color.Blue &&
                    button20.BackColor != Color.Blue &&
                    button21.BackColor != Color.Blue &&
                    button22.BackColor != Color.Blue &&
                    button23.BackColor != Color.Blue &&
                    button24.BackColor != Color.Blue &&
                    button25.BackColor != Color.Blue &&
                    button26.BackColor != Color.Blue &&
                    button27.BackColor != Color.Blue &&
                    button28.BackColor != Color.Blue &&
                    button29.BackColor != Color.Blue &&
                    button30.BackColor != Color.Blue &&
                    button31.BackColor != Color.Blue &&
                    button32.BackColor != Color.Blue &&
                    button33.BackColor != Color.Blue &&
                    button34.BackColor != Color.Blue &&
                    button35.BackColor != Color.Blue &&
                    button36.BackColor != Color.Blue &&
                    button37.BackColor != Color.Blue &&
                    button38.BackColor != Color.Blue &&
                    button39.BackColor != Color.Blue &&
                    button40.BackColor != Color.Blue &&
                    button41.BackColor != Color.Blue &&
                    button42.BackColor != Color.Blue &&
                    button43.BackColor != Color.Blue &&
                    button44.BackColor != Color.Blue &&
                    button45.BackColor != Color.Blue &&
                    button46.BackColor != Color.Blue &&
                    button47.BackColor != Color.Blue &&
                    button48.BackColor != Color.Blue &&
                    button49.BackColor != Color.Blue &&
                    button50.BackColor != Color.Blue)

                {
                    string connectionString = @"" + Properties.Settings.Default.packingConnectionString + "";
                    using (MySqlConnection conn = new MySqlConnection(connectionString))
                    {
                        conn.Open();
                        using (
                            MySqlCommand cmd =
                                new MySqlCommand(
                                    "INSERT INTO dones (project, num_jgut, referencia, dates) VALUES ('" +
                                    Properties.Settings.Default.project + "','" + Properties.Settings.Default.semeistvo +
                                    "','" + Properties.Settings.Default.referencia + "','" +
                                    DateTime.Now + "') ;", conn))
                        {
                            reader = cmd.ExecuteReader();
                            int i = 0;
                            try
                            {
                                while (reader.Read())
                                {
                                    odi = reader[0].ToString();
                                }
                            }
                            catch (Exception)
                            {

                            }
                        }
                    }



                    connectionString = @"" + Settings.Default.packingConnectionString + "";
                    using (MySqlConnection conn = new MySqlConnection(connectionString))
                    {
                        conn.Open();
                        using (
                            MySqlCommand cmd = new MySqlCommand(
                                    "SELECT * FROM dones_smena WHERE num_jgut='"+Properties.Settings.Default.semeistvo+"'" +
                                    "AND project='"+Properties.Settings.Default.project+"' " +
                                    "AND referencia='" + Properties.Settings.Default.referencia +
                                    "' AND dates='" + DateTime.Now.ToString("dd.MM.yyyy") + "';", conn))
                        {
                            reader = cmd.ExecuteReader();
                            int i = 0;
                            while (reader.Read())
                            {
                                odi = reader[0].ToString();
                            }
                            if (odi == "0")
                            {
                                opop();
                            }
                            else
                            {
                                string time = DateTime.Now.ToString("hh");
                                if (Convert.ToInt32(time) < Convert.ToInt32(Properties.Settings.Default.smena))
                                {
                                    connectionString = @"" + Settings.Default.packingConnectionString + "";
                                    using (MySqlConnection conn1 = new MySqlConnection(connectionString))
                                    {
                                        conn1.Open();
                                        using (
                                            MySqlCommand cmd1 =
                                                new MySqlCommand(
                                                    "SELECT smena1 FROM packing.dones_smena WHERE project ='" +
                                                    Properties.Settings.Default.project + "'" +
                                                    " AND num_jgut='" +
                                                    Properties.Settings.Default.semeistvo + "' " +
                                                    "AND referencia='" +
                                                    Properties.Settings.Default.referencia + "' " +
                                                    "AND dates='" + DateTime.Now.ToString("dd.MM.yyyy") +
                                                    "';", conn1))
                                        {
                                            reader1 = cmd1.ExecuteReader();
                                            while (reader1.Read())
                                            {
                                                yan3 = reader1[0].ToString();
                                            }
                                            opop1();
                                        }
                                    }
                                }
                                else
                                {
                                    connectionString = @"" + Settings.Default.packingConnectionString + "";
                                    using (MySqlConnection conn1 = new MySqlConnection(connectionString))
                                    {
                                        conn1.Open();
                                        using (
                                            MySqlCommand cmd1 =
                                                new MySqlCommand(
                                                    "SELECT smena2 FROM packing.dones_smena WHERE project ='" +
                                                    Properties.Settings.Default.project + "'" +
                                                    " AND num_jgut='" +
                                                    Properties.Settings.Default.semeistvo + "' " +
                                                    "AND referencia='" +
                                                    Properties.Settings.Default.referencia + "' " +
                                                    "AND dates='" + DateTime.Now.ToString("dd.MM.yyyy") +
                                                    "';", conn1))
                                        {
                                            reader1 = cmd1.ExecuteReader();
                                            while (reader1.Read())
                                            {
                                                yan3 = reader1[0].ToString();
                                            }
                                            opop1();
                                        }
                                    }
                                }
                            }
                        }

             

                    }
                    HrnsScanForm dbs = new HrnsScanForm();
                    dbs.Show();
                    this.Hide();
                }
                textBox1.Clear();
                
            }
        }

        private void textBox1_KeyUp(object sender, KeyEventArgs e)
        {

        }

        private void opop()
        {
            string connectionString = @"" + Settings.Default.packingConnectionString + "";
            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                conn.Open();
                string time = DateTime.Now.ToString("hh");
                if (Convert.ToInt32(time) < Convert.ToInt32(Properties.Settings.Default.smena))
                {
                    using (
                        MySqlCommand cmd =
                            new MySqlCommand(
                                "INSERT INTO dones_smena (project, num_jgut, referencia, dates, smena1, smena2) VALUES ('" +
                                Properties.Settings.Default.project + "','" + Properties.Settings.Default.semeistvo +
                                "','" + Properties.Settings.Default.referencia + "','" +
                                DateTime.Now.ToString("dd.MM.yyyy") + "', '1', '0' );", conn))
                    {
                        reader = cmd.ExecuteReader();
                        int i = 0;
                        try
                        {
                            while (reader.Read())
                            {
                                odi = reader[0].ToString();
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                }
                else
                {
                    using (
                        MySqlCommand cmd =
                            new MySqlCommand(
                                "INSERT INTO dones_smena (project, num_jgut, referencia, dates, smena1, smena2) VALUES ('" +
                                Properties.Settings.Default.project + "','" + Properties.Settings.Default.semeistvo +
                                "','" + Properties.Settings.Default.referencia + "','" +
                                DateTime.Now.ToString("dd.MM.yyyy") + "', '', '1' );", conn))
                    {
                        reader = cmd.ExecuteReader();
                        int i = 0;
                        try
                        {
                            while (reader.Read())
                            {
                                odi = reader[0].ToString();
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                }
            }
        }

        private void opop1()
        {
            string connectionString = @"" + Settings.Default.packingConnectionString + "";
            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                conn.Open();
                string time = DateTime.Now.ToString("dd");

                int GG = Convert.ToInt32(yan3);
                GG++;
                using (
                    MySqlCommand cmd =
                        new MySqlCommand("UPDATE packing.dones_smena SET smena1 = " + GG + " WHERE project ='" +
                                         Properties.Settings.Default.project + "'" + " AND id='" + odi +
                                         "' AND referencia='" + Properties.Settings.Default.referencia +
                                         "' AND dates='" + DateTime.Now.ToString("dd.MM.yyyy") + "';", conn))
                {
                    cmd.ExecuteNonQuery();
                    reader = cmd.ExecuteReader();
                 
                    int i = 0;
                    //try
                    //{
                    while (reader.Read())
                    {
                        odi = reader[0].ToString();
                    }

                    //}
                    //catch (Exception)
                    //{
                    //}
                }
            }
        }

        private void opop2()
        {
            string connectionString = @"" + Settings.Default.packingConnectionString + "";
            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                conn.Open();
                string time = DateTime.Now.ToString("dd");

                int GG = Convert.ToInt32(yan3);
                GG++;
                using (
                    MySqlCommand cmd =
                        new MySqlCommand("UPDATE packing.dones_smena SET smena2 = " + GG + " WHERE project ='" +
                                         Properties.Settings.Default.project + "'" + " AND id='" + odi +
                                         "' AND referencia='" + Properties.Settings.Default.referencia +
                                         "' AND dates='" + DateTime.Now.ToString("dd.MM.yyyy") + "';", conn))
                {
                    cmd.ExecuteNonQuery();
                    reader = cmd.ExecuteReader();
                    BOXForm dbs = new BOXForm();
                    dbs.Show();
                    this.Hide();
                    int i = 0;
                    //try
                    //{
                    while (reader.Read())
                    {
                        odi = reader[0].ToString();
                    }

                    //}
                    //catch (Exception)
                    //{
                    //}
                }
            }
        }
         
        private void button52_Click(object sender, EventArgs e)
        {
            BOXForm dbs = new BOXForm();
            dbs.Show();
            this.Hide();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}

