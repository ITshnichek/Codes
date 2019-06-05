using MySql.Data.MySqlClient;
using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Drawing.Printing;
using System.IO;

namespace PackingTool
{
    public partial class dbConfig : Form
    {
        
        IniFile INI = new IniFile("config.ini");
        public dbConfig()
        {
            InitializeComponent();
            mtbDBIp.Text = Properties.Settings.Default.DB_IP;
            mtbDBTableName.Text = Properties.Settings.Default.DB_Name;
            mtbDBLogin.Text = Properties.Settings.Default.DB_Login;
            mtbDBPsw.Text = Properties.Settings.Default.DB_Pass;
            tbPrinter2Use.Text = Properties.Settings.Default.LabelPrinter;
            if (Properties.Settings.Default.LabelLType == "EZPL")
            {
                rbEZPL.Checked = true;
            }
            if (Properties.Settings.Default.LabelLType == "ZPL")
            {
                rbZPL.Checked = true;
                gbZPL_Param.Visible = true;
            }
            nmZPL_Start_X.Value = Convert.ToDecimal(Properties.Settings.Default.LabelZPLStart_X);
            nmZPL_Start_Y.Value = Convert.ToDecimal(Properties.Settings.Default.LabelZPLStart_Y);
            nmZPL_Intense.Value = Convert.ToDecimal(Properties.Settings.Default.LabelZPL_Intense);

            tbTimer.Text = Properties.Settings.Default.timer.ToString();
            tbPerfixElTest.Text = Properties.Settings.Default.PerfixElTest;
            tbLenghElTest.Text = Properties.Settings.Default.lenghElTest.ToString();
        }

        private void btDBSaveConf_Click(object sender, EventArgs e)
        {
            string pass = Shifr.Shifrovka(mtbDBPsw.Text, "a8doSuDitOz1hZe#");
            Properties.Settings.Default.DB_IP = mtbDBIp.Text;
            Properties.Settings.Default.DB_Name = mtbDBTableName.Text;
            Properties.Settings.Default.DB_Login = mtbDBLogin.Text;
            Properties.Settings.Default.DB_Pass = pass;
            Properties.Settings.Default.Save();
            
            INI.Write("DB_IP", "DB_IP", mtbDBIp.Text);
            INI.Write("DB_Name", "DB_Name", mtbDBTableName.Text);
            INI.Write("DB_Login", "DB_Login", mtbDBLogin.Text);
            INI.Write("DB_Pass", "DB_Pass", pass);
           // label5.Text = Shifr.DeShifrovka(INI.ReadINI("DB_Pass", "passdb"), "a8doSuDitOz1hZe#");
            MessageBox.Show("Настройки авторизации сохранены", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Asterisk); // Говорим пользователю, что сохранили текст.
        }

        private void btChangeTimer_Click(object sender, EventArgs e)
        {
            if (tbPerfixElTest.Text != "" && tbTimer.Text != "" && tbLenghElTest.Text != "")
            {
                INI.Write("timer", "timer", tbTimer.Text);
                INI.Write("lenghElTest", "lenghElTest", tbLenghElTest.Text);
                INI.Write("PerfixElTest", "PerfixElTest", tbPerfixElTest.Text);
                Properties.Settings.Default.timer = Convert.ToInt32(tbTimer.Text);
                Properties.Settings.Default.lenghElTest = Convert.ToInt32(tbLenghElTest.Text);
                Properties.Settings.Default.PerfixElTest = tbPerfixElTest.Text;

                Properties.Settings.Default.Save();
                lbTimerStatus.Text = "OK";
                lbPerfixStatus.Text = "OK";
                lbLenghStatus.Text = "OK";
            }
            else
            { MessageBox.Show("Заполните все поля"); }

        }

        private void tbTimer_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != 8)
                e.Handled = true;
        }




        private void btDBCheckConf_Click(object sender, EventArgs e)
        {
            string db = "mysql";
            string ConnectDB = @"server=" + mtbDBIp.Text + ";" + "user=" + mtbDBLogin.Text + ";" + "database=" + db + ";" + "password=" + mtbDBPsw.Text + ";charset=utf8;";
            MySqlConnection DBConnection = new MySqlConnection(ConnectDB);
            MySqlCommand command = DBConnection.CreateCommand();
            MySqlDataReader Reader;
            command.CommandText = "SELECT 1";
            try
            {
                DBConnection.Open();
                Reader = command.ExecuteReader();
                if (DBConnection.State == ConnectionState.Open)
                {
                    lbDBStatus.Text = "OK";
                    lbDBStatus.ForeColor = Color.Green;
                }
            }
            catch
            {
                //MessageBox.Show("Нет связи с сервером базы данных!", "Ошибка подключения");
                lbDBStatus.Text = "Error!";
                lbDBStatus.ForeColor = Color.Red;
            }

            finally
            {
                DBConnection.Close();
            }

        }

     

        private void btCreateTable_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Вы действиетльно хотите создать новую базу?", "Выход", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);


            if (result == System.Windows.Forms.DialogResult.OK)
            {
                string pass = Shifr.Shifrovka(mtbDBPsw.Text, "a8doSuDitOz1hZe#");
                if (mtbDBTableName.Text != "" && checkDB() == "0" && pass == INI.ReadINI("DB_Pass", "DB_Pass"))
                {
                    //создание ДБ и таблиц в случае их отсутствия
                    DBConnect sqlcon = new DBConnect();
                    sqlcon.CreateDatabase(mtbDBTableName.Text);
         
                    MessageBox.Show("Таблица " + mtbDBTableName.Text + " успешно создана!");
                }
                else
                {
                    MessageBox.Show("Не указано имя таблицы или она уже существует!");

                }
            }
            else
            {
                Close();
            }

       }

        string checkDB()
        {
            string status = "";
            string db = mtbDBTableName.Text;
            string ConnectDB = @"server=" + mtbDBIp.Text + ";" + "user=" + mtbDBLogin.Text + ";" + "database=" + db + ";" + "password=" + mtbDBPsw.Text + ";charset=utf8;";
            MySqlConnection DBConnection = new MySqlConnection(ConnectDB);
            MySqlCommand command = DBConnection.CreateCommand();
            MySqlDataReader Reader;
            command.CommandText = "SELECT 1";
            try
            {
                DBConnection.Open();
                Reader = command.ExecuteReader();
                if (DBConnection.State == ConnectionState.Open)
                {
                    status = "1";
                   // return status;
                }
            }
            catch
            {
                status = "0";
               // return status;
            }

            finally
            {
                
                DBConnection.Close();
            }
            return status;
        }
        private void btPrinterSelect_Click(object sender, EventArgs e)
        {
            PrintDialog pd = new PrintDialog();
            pd.PrinterSettings = new PrinterSettings();
            if (DialogResult.OK == pd.ShowDialog(this))
            {
                INI.Write("LabelPrinter", "LabelPrinter", pd.PrinterSettings.PrinterName);
                Properties.Settings.Default.LabelPrinter = pd.PrinterSettings.PrinterName;
                Properties.Settings.Default.Save();
                tbPrinter2Use.Text = INI.ReadINI("LabelPrinter", "LabelPrinter");
            }
        }

        private void rbZPL_CheckedChanged(object sender, EventArgs e)
        {
            if (rbZPL.Checked == true)
            {
                INI.Write("LabelLType", "LabelLType", "ZPL");
                Properties.Settings.Default.LabelLType = "ZPL";
                Properties.Settings.Default.Save();
                gbZPL_Param.Visible = true;
            }
        }

        private void rbEZPL_CheckedChanged(object sender, EventArgs e)
        {
            if (rbEZPL.Checked == true)
            {
                INI.Write("LabelLType", "LabelLType", "EZPL");
                Properties.Settings.Default.LabelLType = "EZPL";
                Properties.Settings.Default.Save();
                gbZPL_Param.Visible = false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Label_Printing.ZPL_fin_label("test", "test", "12.34.5678", "12:34", 1);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            INI.Write("LabelZPLStart_X", "LabelZPLStart_X", (Convert.ToInt32(nmZPL_Start_X.Value)).ToString());
            INI.Write("LabelZPLStart_Y", "LabelZPLStart_Y", (Convert.ToInt32(nmZPL_Start_Y.Value)).ToString());
            INI.Write("LabelZPL_Intense", "LabelZPL_Intense", (Convert.ToInt32(nmZPL_Intense.Value)).ToString());
            Properties.Settings.Default.LabelZPLStart_X = Convert.ToInt32(nmZPL_Start_X.Value);
            Properties.Settings.Default.LabelZPLStart_Y = Convert.ToInt32(nmZPL_Start_Y.Value);
            Properties.Settings.Default.LabelZPL_Intense = Convert.ToInt32(nmZPL_Intense.Value);
            Properties.Settings.Default.Save();
        }

        private void rbZPL_CheckedChanged_1(object sender, EventArgs e)
        {
            if (rbZPL.Checked == true)
            {
                INI.Write("LabelLType", "LabelLType", "ZPL");
                Properties.Settings.Default.LabelLType = "ZPL";
                Properties.Settings.Default.Save();
                gbZPL_Param.Visible = true;
            }
        }

        private void rbEZPL_CheckedChanged_1(object sender, EventArgs e)
        {
            if (rbEZPL.Checked == true)
            {
                INI.Write("LabelLType", "LabelLType", "EZPL");
                Properties.Settings.Default.LabelLType = "EZPL";
                Properties.Settings.Default.Save();
                gbZPL_Param.Visible = false;
            }
        }

        private void dbConfig_Load(object sender, EventArgs e)
        {
            dateTimePicker1.Value = DateTime.Today;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try {
            string ololo = Convert.ToString(DateTime.Today);
            string[] go = ololo.Split('.', ' ');
            string constring = "
            string file = @"" + go[0] + "-" + go[1] + "-" + go[2] + ".sql";
            using (MySqlConnection conn = new MySqlConnection(constring))
            {
                using (MySqlCommand cmd = new MySqlCommand())
                {
                    using (MySqlBackup mb = new MySqlBackup(cmd))
                    {
                        cmd.Connection = conn;
                        conn.Open();
                        mb.ExportToFile(file);
                        conn.Close();
                        MessageBox.Show("Резервная копия создана!");
                    }
                }
            }
        }
            catch (Exception)
            {
                MessageBox.Show("Произошла ошибка в создании резервной копии!");
            }
}

        private void button2_Click_1(object sender, EventArgs e)
        {
            try
            {
                string ololo = Convert.ToString(dateTimePicker1.Value);
                string[] go = ololo.Split('.', ' ');
                string constring = 
                string file = @"" + go[0] + "-" + go[1] + "-" + go[2] + ".sql";
                using (MySqlConnection conn = new MySqlConnection(constring))
                {
                    using (MySqlCommand cmd = new MySqlCommand())
                    {
                        using (MySqlBackup mb = new MySqlBackup(cmd))
                        {
                            cmd.Connection = conn;
                            conn.Open();
                            mb.ImportFromFile(file);
                            conn.Close();
                            MessageBox.Show("Резервная копия восстановленна!");
                        }
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Данной резервной копии не существует!");
            }
        }

        private void dateTimePicker1_KeyDown(object sender, KeyEventArgs e)
        {
           
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }
    }
}
