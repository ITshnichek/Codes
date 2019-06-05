using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing.Printing;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace PrintSigip
{
    public partial class Form1 : Form
    {
        string BT;
        string BT2;
        string BT3;
        string Zachistka;
        string Zachistka2;
        string Zachistka3;
        string login1;
        string Two1;
        string Three1;
        string Four1;
        string Five1;
        string Six1;
        string Nine1;
        string Ten1;
        string Two2;
        string Three2;
        string Four2;
        string Five2;
        string Six2;
        string Nine2;
        string Ten2;
        string Two3;
        string Three3;
        string Four3;
        string Five3;
        string Six3;
        string Nine3;
        string Ten3;
        string Listt;
        int k = 0;
        
        public Form1()
        {
            InitializeComponent();
        }
        string[] Ghost;
        string files;
        string arr;
        public static System.Drawing.Brush WindowFrame { get; }
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog fDialog = new OpenFileDialog();
                fDialog.Title = "Open Files";
                fDialog.Filter = "xls files(*.xls)|*.xls*";
                fDialog.InitialDirectory = @"C:\";
                if (fDialog.ShowDialog() == DialogResult.OK)
                {
                    arr = fDialog.FileName.ToString();
                }
                Ghost = arr.Split(new char[] { '\\' });
                int ghos = Ghost.Length;
                label1.Text = Ghost[ghos - 1];
            }
            catch (Exception)
            {
                MessageBox.Show("Не выбран файл");
            }
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {

        }
        string[] wordsv2 = null;
        string HHv2 = null;
        private void button2_Click(object sender, EventArgs e)
        {
            Zachistka = textBox7.Text;
            Zachistka2 = textBox8.Text;
            Zachistka3 = textBox9.Text;

////Некоторая информация была удалена во избежании утечки конфидециальных данных


            login1 = SystemInformation.UserName;
            Dictionary<string, int> Txns = new Dictionary<string, int>();
            Txns.Clear();
            Txns = null;
            GC.Collect();
            int p = 0;
            string ConStr = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + arr + ";" + "Extended Properties=\"Excel 8.0; HDR=NO; IMEX=1\"";
            System.Data.DataSet ds = new System.Data.DataSet("EXCEL");
            OleDbConnection cn = new OleDbConnection(ConStr);
            try
            {
                cn.Open();
            }
            catch (Exception)
            {
                MessageBox.Show("Проверьте правильность выбора файла и повторите попытку");
                return;
            }
            System.Data.DataTable schemaTable = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            string sheet1;
            try
            {
                for (int i2 = 0; ; i2++)
                {
                    sheet1 = (string)schemaTable.Rows[i2].ItemArray[2];
                }
            }
            catch (Exception)
            {
            }
            sheet1 = (string)schemaTable.Rows[p].ItemArray[2];
            string select = String.Format("SELECT * FROM [{0}]", sheet1);
            OleDbDataAdapter ad = new OleDbDataAdapter(select, cn);
            ad.Fill(ds);
            System.Data.DataTable tb = ds.Tables[0];
            int kek= 1;
            int l = 0;
        m1:
            if (l>2000)
            {
                MessageBox.Show("Проверьте правильность данных и повторите попытку");
                return;
            }
            try
            {
                for (int i = 0; ; i++)
                {
                    HHv2 = HHv2 + ";" + Convert.ToString(tb.Rows[l].ItemArray[i]);

                }
            }
            catch (Exception)
            {
                try
                {
                    wordsv2 = HHv2.Split(new char[] { ';' });
                }
                catch (Exception)
                {
                }
            }

            if (kek==1 && (textBox1.Text != null && textBox1.Text != ""))
            {
                if (wordsv2[5] == textBox1.Text && (Two1 == null || Two1 == "") && kek == 1)
                {
                    Two1 = wordsv2[2];
                    Three1 = wordsv2[3];
                    Four1 = wordsv2[4];
                    Five1 = wordsv2[5];
                    Six1 = wordsv2[6];
                    Nine1 = wordsv2[9];
                    Ten1 = wordsv2[10];
                    l = 0;
                    kek++;
                }
                else
                {
                    HHv2 = null;
                    l++;
                    goto m1;
                }
            }

            if (kek == 2 && (textBox2.Text != null && textBox2.Text != ""))
            {
                if (wordsv2[5] == textBox2.Text && (Two2 == null || Two2 == "") && kek ==2)
                {
                    Two2 = wordsv2[2];
                    Three2 = wordsv2[3];
                    Four2 = wordsv2[4];
                    Five2 = wordsv2[5];
                    Six2 = wordsv2[6];
                    Nine2 = wordsv2[9];
                    Ten2 = wordsv2[10];
                    l = 0;
                    kek++;
                }
                else
                {
                    HHv2 = null;
                    l++;
                    goto m1;
                }
            }

            if (kek == 3 && (textBox3.Text != null && textBox3.Text != ""))
            {
                if (wordsv2[5] == textBox3.Text && (Two3 == null || Two3 == "") && kek ==3)
                {
                    Two3 = wordsv2[2];
                    Three3 = wordsv2[3];
                    Four3 = wordsv2[4];
                    Five3 = wordsv2[5];
                    Six3 = wordsv2[6];
                    Nine3 = wordsv2[9];
                    Ten3 = wordsv2[10];
                    kek++;
                }
                else
                {
                    HHv2 = null;
                    l++;
                    goto m1;
                }
            }
            
            PRD();

            Excel.Application excelApp = new Excel.Application();

            Excel.Workbook wb = excelApp.Workbooks.Open(@"H:\PrintSJ.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets[k];
            ws.PageSetup.PrintHeadings = false;
            ws.PageSetup.BlackAndWhite = false;
            ws.PageSetup.PrintGridlines = false;
            ws.PageSetup.PrintTitleRows = "$1:$2"; 
            ws.PrintOut(Type.Missing, Type.Missing, Type.Missing, Type.Missing,Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.FinalReleaseComObject(ws);

            wb.Close(false, Type.Missing, Type.Missing);
            Marshal.FinalReleaseComObject(wb);

            excelApp.Quit();
            Marshal.FinalReleaseComObject(excelApp);
            System.IO.File.Delete(@"H:\PrintSJ.xlsx");
            System.Windows.Forms.Application.Restart();
            return;
        }

        public void PRD()
        {
            if (textBox1.Text!=null && textBox2.Text == "" && textBox3.Text == "")
            {
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWb = xlApp.Workbooks.Open(@"H:\PJ\Excel.xlsx");
                Excel.Worksheet xlSht = xlWb.Sheets[1];
                int iLastRow = xlSht.Cells[xlSht.Rows.Count, "H"].End[Excel.XlDirection.xlUp].Row;

                
                xlSht.Cells[1, "C"].Value = Three1.ToString();
                xlSht.Cells[1, "D"].Value = Four1.ToString();
                xlSht.Cells[2, "D"].Value = Zachistka.ToString();
                xlSht.Cells[3, "B"].Value = Five1.ToString();
                xlSht.Cells[4, "B"].Value = Six1.ToString();
                xlSht.Cells[6, "B"].Value = Nine1.ToString();
                xlSht.Cells[8, "B"].Value = BT.ToString();
                Colors(xlSht);
                if (Four1=="S2" || Four1=="С2" || Four1=="C2" || Four1=="A2" || Four1=="a2" || Four1=="А2" || Four1=="а2")
                {
                    (xlSht.Cells[1, "D"] as Range).Font.Color = Color.Black;
                    (xlSht.Cells[1, "H"] as Range).Font.Color = Color.Black;
                }

                xlSht.Cells[6, "D"].Value = Ten1.ToString();
                xlSht.Cells[7, "C"].Value = Two1.ToString();

                xlSht.Cells[1, "G"].Value = Three1.ToString();
                xlSht.Cells[1, "H"].Value = Four1.ToString();
                xlSht.Cells[2, "H"].Value = Zachistka.ToString();
                xlSht.Cells[3, "F"].Value = Five1.ToString();
                xlSht.Cells[4, "F"].Value = Six1.ToString();
                xlSht.Cells[6, "F"].Value = Nine1.ToString();
                xlSht.Cells[6, "H"].Value = Ten1.ToString();
                xlSht.Cells[7, "G"].Value = Two1.ToString();
                xlSht.Cells[8, "F"].Value = BT.ToString();
                Listt = "$1";
                xlWb.SaveAs(@"H:\PrintSJ.xlsx");
                xlWb.Close();

                xlApp.Quit(); ;
                k = 1;
            }
            else if (textBox1.Text != null && textBox2.Text != "" && textBox3.Text == "")
            {
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWb = xlApp.Workbooks.Open(@"H:\PJ\Excel.xlsx");
                Excel.Worksheet xlSht = xlWb.Sheets[2];
                int iLastRow = xlSht.Cells[xlSht.Rows.Count, "H"].End[Excel.XlDirection.xlUp].Row;
                ///1 para
                xlSht.Cells[1, "C"].Value = Three1.ToString();
                xlSht.Cells[1, "D"].Value = Four1.ToString();
                xlSht.Cells[2, "D"].Value = Zachistka.ToString();
                xlSht.Cells[3, "B"].Value = Five1.ToString();
                xlSht.Cells[4, "B"].Value = Six1.ToString();
                xlSht.Cells[6, "B"].Value = Nine1.ToString();
                xlSht.Cells[6, "D"].Value = Ten1.ToString();
                xlSht.Cells[7, "C"].Value = Two1.ToString();
                xlSht.Cells[8, "B"].Value = BT.ToString();
                Colors(xlSht);
                if (Four1 == "S2" || Four1 == "С2" || Four1 == "C2" || Four1 == "A2" || Four1 == "a2" || Four1 == "А2" || Four1 == "а2")
                {
                    (xlSht.Cells[1, "D"] as Range).Font.Color = Color.Black;
                    (xlSht.Cells[1, "H"] as Range).Font.Color = Color.Black;
                }

                xlSht.Cells[1, "G"].Value = Three1.ToString();
                xlSht.Cells[1, "H"].Value = Four1.ToString();
                xlSht.Cells[2, "H"].Value = Zachistka.ToString();
                xlSht.Cells[3, "F"].Value = Five1.ToString();
                xlSht.Cells[4, "F"].Value = Six1.ToString();
                xlSht.Cells[6, "F"].Value = Nine1.ToString();
                xlSht.Cells[6, "H"].Value = Ten1.ToString();
                xlSht.Cells[7, "G"].Value = Two1.ToString();
                xlSht.Cells[8, "F"].Value = BT.ToString();

                ///2para
                xlSht.Cells[9, "C"].Value = Three2.ToString();
                xlSht.Cells[10, "D"].Value = Zachistka2.ToString();
                xlSht.Cells[9, "D"].Value = Four2.ToString();
                xlSht.Cells[11, "B"].Value = Five2.ToString();
                xlSht.Cells[12, "B"].Value = Six2.ToString();
                xlSht.Cells[14, "B"].Value = Nine2.ToString();
                xlSht.Cells[14, "D"].Value = Ten2.ToString();
                xlSht.Cells[15, "C"].Value = Two2.ToString();
                xlSht.Cells[16, "B"].Value = BT2.ToString();
                Colors2(xlSht);
                if (Four2 == "S2" || Four2 == "С2" || Four2 == "C2" || Four2 == "A2" || Four2 == "a2" || Four2 == "А2" || Four2 == "а2")
                {
                    (xlSht.Cells[9, "D"] as Range).Font.Color = Color.Black;
                    (xlSht.Cells[9, "H"] as Range).Font.Color = Color.Black;
                }
                xlSht.Cells[9, "G"].Value = Three2.ToString();
                xlSht.Cells[9, "H"].Value = Four2.ToString();
                xlSht.Cells[10, "H"].Value = Zachistka2.ToString();
                xlSht.Cells[11, "F"].Value = Five2.ToString();
                xlSht.Cells[12, "F"].Value = Six2.ToString();
                xlSht.Cells[14, "F"].Value = Nine2.ToString();
                xlSht.Cells[14, "H"].Value = Ten2.ToString();
                xlSht.Cells[15, "G"].Value = Two2.ToString();
                xlSht.Cells[16, "F"].Value = BT2.ToString();
                Listt = "$2";
                xlWb.SaveAs(@"H:\PrintSJ.xlsx");
                xlWb.Close();

                xlApp.Quit(); ;
                k = 2;
            }
            else if (textBox1.Text != null && textBox2.Text != "" && textBox3.Text != "")
            {
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWb = xlApp.Workbooks.Open(@"H:\PJ\Excel.xlsx");
                Excel.Worksheet xlSht = xlWb.Sheets[3];
                int iLastRow = xlSht.Cells[xlSht.Rows.Count, "I"].End[Excel.XlDirection.xlUp].Row;
                ////1Para
                xlSht.Cells[1, "C"].Value = Three1.ToString();
                xlSht.Cells[1, "D"].Value = Four1.ToString();
                xlSht.Cells[2, "D"].Value = Zachistka.ToString();
                xlSht.Cells[3, "B"].Value = Five1.ToString();
                xlSht.Cells[4, "B"].Value = Six1.ToString();
                xlSht.Cells[6, "B"].Value = Nine1.ToString();
                xlSht.Cells[6, "D"].Value = Ten1.ToString();
                xlSht.Cells[7, "C"].Value = Two1.ToString();
                xlSht.Cells[8, "B"].Value = BT.ToString();
                Colors(xlSht);
                if (Four1 == "S2" || Four1 == "С2" || Four1 == "C2" || Four1 == "A2" || Four1 == "a2" || Four1 == "А2" || Four1 == "а2")
                {
                    (xlSht.Cells[1, "D"] as Range).Font.Color = Color.Black;
                    (xlSht.Cells[1, "H"] as Range).Font.Color = Color.Black;
                }
                xlSht.Cells[1, "G"].Value = Three1.ToString();
                xlSht.Cells[1, "H"].Value = Four1.ToString();
                xlSht.Cells[2, "H"].Value = Zachistka.ToString();
                xlSht.Cells[3, "F"].Value = Five1.ToString();
                xlSht.Cells[4, "F"].Value = Six1.ToString();
                xlSht.Cells[6, "F"].Value = Nine1.ToString();
                xlSht.Cells[6, "H"].Value = Ten1.ToString();
                xlSht.Cells[7, "G"].Value = Two1.ToString();
                xlSht.Cells[8, "F"].Value = BT.ToString();

                ///2Para
                xlSht.Cells[9, "C"].Value = Three2.ToString();
                xlSht.Cells[9, "D"].Value = Four2.ToString();
                xlSht.Cells[10, "D"].Value = Zachistka2.ToString();
                xlSht.Cells[11, "B"].Value = Five2.ToString();
                xlSht.Cells[12, "B"].Value = Six2.ToString();
                xlSht.Cells[14, "B"].Value = Nine2.ToString();
                xlSht.Cells[14, "D"].Value = Ten2.ToString();
                xlSht.Cells[15, "C"].Value = Two2.ToString();
                xlSht.Cells[16, "B"].Value = BT2.ToString();
                Colors2(xlSht);
                if (Four2 == "S2" || Four2 == "С2" || Four2 == "C2" || Four2 == "A2" || Four2 == "a2" || Four2 == "А2" || Four2 == "а2")
                {
                    (xlSht.Cells[9, "D"] as Range).Font.Color = Color.Black;
                    (xlSht.Cells[9, "H"] as Range).Font.Color = Color.Black;
                }
                xlSht.Cells[9, "G"].Value = Three2.ToString();
                xlSht.Cells[9, "H"].Value = Four2.ToString();
                xlSht.Cells[10, "H"].Value = Zachistka2.ToString();
                xlSht.Cells[11, "F"].Value = Five2.ToString();
                xlSht.Cells[12, "F"].Value = Six2.ToString();
                xlSht.Cells[14, "F"].Value = Nine2.ToString();
                xlSht.Cells[14, "H"].Value = Ten2.ToString();
                xlSht.Cells[15, "G"].Value = Two2.ToString();
                xlSht.Cells[16, "F"].Value = BT2.ToString();

                ////3Para
                xlSht.Cells[17, "C"].Value = Three3.ToString();
                xlSht.Cells[17, "D"].Value = Four3.ToString();
                xlSht.Cells[18, "D"].Value = Zachistka3.ToString();
                xlSht.Cells[19, "B"].Value = Five3.ToString();
                xlSht.Cells[20, "B"].Value = Six3.ToString();
                xlSht.Cells[22, "B"].Value = Nine3.ToString();
                xlSht.Cells[22, "D"].Value = Ten3.ToString();
                xlSht.Cells[23, "C"].Value = Two3.ToString();
                xlSht.Cells[24, "B"].Value = Two1.ToString();
                xlSht.Cells[15, "D"].Value = "".ToString();
                xlSht.Cells[23, "D"].Value = "".ToString();
                xlSht.Cells[23, "H"].Value = "".ToString();
                xlSht.Cells[24, "B"].Value = BT3.ToString();
                Colors3(xlSht);
                if (Four3 == "S2" || Four3 == "С2" || Four3 == "C2" || Four3 == "A2" || Four3 == "a2" || Four3 == "А2" || Four3 == "а2")
                {
                    (xlSht.Cells[17, "D"] as Range).Font.Color = Color.Black;
                    (xlSht.Cells[17, "H"] as Range).Font.Color = Color.Black;
                }
                xlSht.Cells[17, "G"].Value = Three3.ToString();
                xlSht.Cells[17, "H"].Value = Four3.ToString();
                xlSht.Cells[18, "H"].Value = Zachistka3.ToString();
                xlSht.Cells[19, "F"].Value = Five3.ToString();
                xlSht.Cells[20, "F"].Value = Six3.ToString();
                xlSht.Cells[22, "F"].Value = Nine3.ToString();
                xlSht.Cells[22, "H"].Value = Ten3.ToString();
                xlSht.Cells[23, "G"].Value = Two3.ToString();
                xlSht.Cells[24, "F"].Value = BT3.ToString();
                Listt = "$3";
                xlWb.SaveAs(@"H:\PrintSJ.xlsx");
                xlWb.Close();
                
                xlApp.Quit(); ;
                k = 3;
            }

        }

        public void Colors(Excel.Worksheet xlSht)
        {
            if (Six1 == "Синий" || Six1 == "синий")
            {
                (xlSht.Cells[4, "B"] as Range).Interior.Color = Color.Blue;
                (xlSht.Cells[4, "B"] as Range).Font.Color = Color.White;
                (xlSht.Cells[4, "C"] as Range).Interior.Color = Color.Blue;
                (xlSht.Cells[4, "D"] as Range).Interior.Color = Color.Blue;

                (xlSht.Cells[4, "F"] as Range).Interior.Color = Color.Blue;
                (xlSht.Cells[4, "F"] as Range).Font.Color = Color.White;
                (xlSht.Cells[4, "G"] as Range).Interior.Color = Color.Blue;
                (xlSht.Cells[4, "H"] as Range).Interior.Color = Color.Blue;
            }
            else if (Six1 == "Желтый" || Six1 == "желтый")
            {
                (xlSht.Cells[4, "B"] as Range).Interior.Color = Color.Yellow;
                (xlSht.Cells[4, "C"] as Range).Interior.Color = Color.Yellow;
                (xlSht.Cells[4, "D"] as Range).Interior.Color = Color.Yellow;

                (xlSht.Cells[4, "F"] as Range).Interior.Color = Color.Yellow;
                (xlSht.Cells[4, "G"] as Range).Interior.Color = Color.Yellow;
                (xlSht.Cells[4, "H"] as Range).Interior.Color = Color.Yellow;
            }
            else if (Six1 == "Коричневый" || Six1 == "коричневый")
            {
                (xlSht.Cells[4, "B"] as Range).Interior.Color = Color.Brown;
                (xlSht.Cells[4, "B"] as Range).Font.Color = Color.White;
                (xlSht.Cells[4, "C"] as Range).Interior.Color = Color.Brown;
                (xlSht.Cells[4, "D"] as Range).Interior.Color = Color.Brown;

                (xlSht.Cells[4, "F"] as Range).Interior.Color = Color.Brown;
                (xlSht.Cells[4, "F"] as Range).Font.Color = Color.White;
                (xlSht.Cells[4, "G"] as Range).Interior.Color = Color.Brown;
                (xlSht.Cells[4, "H"] as Range).Interior.Color = Color.Brown;
            }
            else if (Six1 == "Черный" || Six1 == "черный")
            {
                (xlSht.Cells[4, "B"] as Range).Interior.Color = Color.Black;
                (xlSht.Cells[4, "B"] as Range).Font.Color = Color.White;
                (xlSht.Cells[4, "C"] as Range).Interior.Color = Color.Black;
                (xlSht.Cells[4, "D"] as Range).Interior.Color = Color.Black;

                (xlSht.Cells[4, "F"] as Range).Interior.Color = Color.Black;
                (xlSht.Cells[4, "F"] as Range).Font.Color = Color.White;
                (xlSht.Cells[4, "G"] as Range).Interior.Color = Color.Black;
                (xlSht.Cells[4, "H"] as Range).Interior.Color = Color.Black;
            }
            else if (Six1 == "Красный" || Six1 == "красный")
            {
                (xlSht.Cells[4, "B"] as Range).Interior.Color = Color.Red;
                (xlSht.Cells[4, "B"] as Range).Font.Color = Color.White;
                (xlSht.Cells[4, "C"] as Range).Interior.Color = Color.Red;
                (xlSht.Cells[4, "D"] as Range).Interior.Color = Color.Red;

                (xlSht.Cells[4, "F"] as Range).Interior.Color = Color.Red;
                (xlSht.Cells[4, "F"] as Range).Font.Color = Color.White;
                (xlSht.Cells[4, "G"] as Range).Interior.Color = Color.Red;
                (xlSht.Cells[4, "H"] as Range).Interior.Color = Color.Red;
            }
            else if (Six1 == "Зеленый" || Six1 == "зеленый")
            {
                (xlSht.Cells[4, "B"] as Range).Interior.Color = Color.Green;
                (xlSht.Cells[4, "B"] as Range).Font.Color = Color.White;
                (xlSht.Cells[4, "C"] as Range).Interior.Color = Color.Green;
                (xlSht.Cells[4, "D"] as Range).Interior.Color = Color.Green;

                (xlSht.Cells[4, "F"] as Range).Interior.Color = Color.Green;
                (xlSht.Cells[4, "F"] as Range).Font.Color = Color.White;
                (xlSht.Cells[4, "G"] as Range).Interior.Color = Color.Green;
                (xlSht.Cells[4, "H"] as Range).Interior.Color = Color.Green;
            }
            else if (Six1 == "Серый" || Six1 == "серый")
            {
                (xlSht.Cells[4, "B"] as Range).Interior.Color = Color.Gray;
                (xlSht.Cells[4, "C"] as Range).Interior.Color = Color.Gray;
                (xlSht.Cells[4, "D"] as Range).Interior.Color = Color.Gray;

                (xlSht.Cells[4, "F"] as Range).Interior.Color = Color.Gray;
                (xlSht.Cells[4, "G"] as Range).Interior.Color = Color.Gray;
                (xlSht.Cells[4, "H"] as Range).Interior.Color = Color.Gray;
            }
            else if (Six1 == "Оранжевый" || Six1 == "оранжевый")
            {
                (xlSht.Cells[4, "B"] as Range).Interior.Color = Color.Orange;
                (xlSht.Cells[4, "C"] as Range).Interior.Color = Color.Orange;
                (xlSht.Cells[4, "D"] as Range).Interior.Color = Color.Orange;

                (xlSht.Cells[4, "F"] as Range).Interior.Color = Color.Orange;
                (xlSht.Cells[4, "G"] as Range).Interior.Color = Color.Orange;
                (xlSht.Cells[4, "H"] as Range).Interior.Color = Color.Orange;
            }
            else if (Six1 == "Розовый" || Six1 == "розовый")
            {
                (xlSht.Cells[4, "B"] as Range).Interior.Color = Color.Pink;
                (xlSht.Cells[4, "C"] as Range).Interior.Color = Color.Pink;
                (xlSht.Cells[4, "D"] as Range).Interior.Color = Color.Pink;

                (xlSht.Cells[4, "F"] as Range).Interior.Color = Color.Pink;
                (xlSht.Cells[4, "G"] as Range).Interior.Color = Color.Pink;
                (xlSht.Cells[4, "H"] as Range).Interior.Color = Color.Pink;
            }
            else if (Six1 == "Фиолетовый" || Six1 == "фиолетовый")
            {
                (xlSht.Cells[4, "B"] as Range).Interior.Color = Color.Violet;
                (xlSht.Cells[4, "B"] as Range).Font.Color = Color.White;
                (xlSht.Cells[4, "C"] as Range).Interior.Color = Color.Violet;
                (xlSht.Cells[4, "D"] as Range).Interior.Color = Color.Violet;

                (xlSht.Cells[4, "F"] as Range).Interior.Color = Color.Violet;
                (xlSht.Cells[4, "F"] as Range).Font.Color = Color.White;
                (xlSht.Cells[4, "G"] as Range).Interior.Color = Color.Violet;
                (xlSht.Cells[4, "H"] as Range).Interior.Color = Color.Violet;
            }
        }
        public void Colors2(Excel.Worksheet xlSht)
        {
            if (Six2 == "Синий" || Six2 == "синий")
            {
                (xlSht.Cells[12, "B"] as Range).Interior.Color = Color.Blue;
                (xlSht.Cells[12, "B"] as Range).Font.Color = Color.White;
                (xlSht.Cells[12, "C"] as Range).Interior.Color = Color.Blue;
                (xlSht.Cells[12, "D"] as Range).Interior.Color = Color.Blue;

                (xlSht.Cells[12, "F"] as Range).Interior.Color = Color.Blue;
                (xlSht.Cells[12, "F"] as Range).Font.Color = Color.White;
                (xlSht.Cells[12, "G"] as Range).Interior.Color = Color.Blue;
                (xlSht.Cells[12, "H"] as Range).Interior.Color = Color.Blue;
            }
            else if (Six2 == "Желтый" || Six2 == "желтый")
            {
                (xlSht.Cells[12, "B"] as Range).Interior.Color = Color.Yellow;
                (xlSht.Cells[12, "C"] as Range).Interior.Color = Color.Yellow;
                (xlSht.Cells[12, "D"] as Range).Interior.Color = Color.Yellow;

                (xlSht.Cells[12, "F"] as Range).Interior.Color = Color.Yellow;
                (xlSht.Cells[12, "G"] as Range).Interior.Color = Color.Yellow;
                (xlSht.Cells[12, "H"] as Range).Interior.Color = Color.Yellow;
            }
            else if (Six2 == "Коричневый" || Six2 == "коричневый")
            {
                (xlSht.Cells[12, "B"] as Range).Interior.Color = Color.Brown;
                (xlSht.Cells[12, "B"] as Range).Font.Color = Color.White;
                (xlSht.Cells[12, "C"] as Range).Interior.Color = Color.Brown;
                (xlSht.Cells[12, "D"] as Range).Interior.Color = Color.Brown;

                (xlSht.Cells[12, "F"] as Range).Interior.Color = Color.Brown;
                (xlSht.Cells[12, "F"] as Range).Font.Color = Color.White;
                (xlSht.Cells[12, "G"] as Range).Interior.Color = Color.Brown;
                (xlSht.Cells[12, "H"] as Range).Interior.Color = Color.Brown;
            }
            else if (Six2 == "Черный" || Six2 == "черный")
            {
                (xlSht.Cells[12, "B"] as Range).Interior.Color = Color.Black;
                (xlSht.Cells[12, "B"] as Range).Font.Color = Color.White;
                (xlSht.Cells[12, "C"] as Range).Interior.Color = Color.Black;
                (xlSht.Cells[12, "D"] as Range).Interior.Color = Color.Black;

                (xlSht.Cells[12, "F"] as Range).Interior.Color = Color.Black;
                (xlSht.Cells[12, "F"] as Range).Font.Color = Color.White;
                (xlSht.Cells[12, "G"] as Range).Interior.Color = Color.Black;
                (xlSht.Cells[12, "H"] as Range).Interior.Color = Color.Black;
            }
            else if (Six2 == "Красный" || Six2 == "красный")
            {
                (xlSht.Cells[12, "B"] as Range).Interior.Color = Color.Red;
                (xlSht.Cells[12, "B"] as Range).Font.Color = Color.White;
                (xlSht.Cells[12, "C"] as Range).Interior.Color = Color.Red;
                (xlSht.Cells[12, "D"] as Range).Interior.Color = Color.Red;

                (xlSht.Cells[12, "F"] as Range).Interior.Color = Color.Red;
                (xlSht.Cells[12, "F"] as Range).Font.Color = Color.White;
                (xlSht.Cells[12, "G"] as Range).Interior.Color = Color.Red;
                (xlSht.Cells[12, "H"] as Range).Interior.Color = Color.Red;
            }
            else if (Six2 == "Зеленый" || Six2 == "зеленый")
            {
                (xlSht.Cells[12, "B"] as Range).Interior.Color = Color.Green;
                (xlSht.Cells[12, "B"] as Range).Font.Color = Color.White;
                (xlSht.Cells[12, "C"] as Range).Interior.Color = Color.Green;
                (xlSht.Cells[12, "D"] as Range).Interior.Color = Color.Green;

                (xlSht.Cells[12, "F"] as Range).Interior.Color = Color.Green;
                (xlSht.Cells[12, "F"] as Range).Font.Color = Color.White;
                (xlSht.Cells[12, "G"] as Range).Interior.Color = Color.Green;
                (xlSht.Cells[12, "H"] as Range).Interior.Color = Color.Green;
            }
            else if (Six2 == "Серый" || Six2 == "серый")
            {
                (xlSht.Cells[12, "B"] as Range).Interior.Color = Color.Gray;
                (xlSht.Cells[12, "C"] as Range).Interior.Color = Color.Gray;
                (xlSht.Cells[12, "D"] as Range).Interior.Color = Color.Gray;

                (xlSht.Cells[12, "F"] as Range).Interior.Color = Color.Gray;
                (xlSht.Cells[12, "G"] as Range).Interior.Color = Color.Gray;
                (xlSht.Cells[12, "H"] as Range).Interior.Color = Color.Gray;
            }
            else if (Six2 == "Оранжевый" || Six2 == "оранжевый")
            {
                (xlSht.Cells[12, "B"] as Range).Interior.Color = Color.Orange;
                (xlSht.Cells[12, "C"] as Range).Interior.Color = Color.Orange;
                (xlSht.Cells[12, "D"] as Range).Interior.Color = Color.Orange;

                (xlSht.Cells[12, "F"] as Range).Interior.Color = Color.Orange;
                (xlSht.Cells[12, "G"] as Range).Interior.Color = Color.Orange;
                (xlSht.Cells[12, "H"] as Range).Interior.Color = Color.Orange;
            }
            else if (Six2 == "Розовый" || Six2 == "розовый")
            {
                (xlSht.Cells[12, "B"] as Range).Interior.Color = Color.Pink;
                (xlSht.Cells[12, "C"] as Range).Interior.Color = Color.Pink;
                (xlSht.Cells[12, "D"] as Range).Interior.Color = Color.Pink;

                (xlSht.Cells[12, "F"] as Range).Interior.Color = Color.Pink;
                (xlSht.Cells[12, "G"] as Range).Interior.Color = Color.Pink;
                (xlSht.Cells[12, "H"] as Range).Interior.Color = Color.Pink;
            }
            else if (Six2 == "Фиолетовый" || Six2 == "фиолетовый")
            {
                (xlSht.Cells[12, "B"] as Range).Interior.Color = Color.Violet;
                (xlSht.Cells[12, "B"] as Range).Font.Color = Color.White;
                (xlSht.Cells[12, "C"] as Range).Interior.Color = Color.Violet;
                (xlSht.Cells[12, "D"] as Range).Interior.Color = Color.Violet;

                (xlSht.Cells[12, "F"] as Range).Interior.Color = Color.Violet;
                (xlSht.Cells[12, "F"] as Range).Font.Color = Color.White;
                (xlSht.Cells[12, "G"] as Range).Interior.Color = Color.Violet;
                (xlSht.Cells[12, "H"] as Range).Interior.Color = Color.Violet;
            }
        }
        public void Colors3(Excel.Worksheet xlSht)
        {
            if (Six3 == "Синий" || Six3 == "синий")
            {
                (xlSht.Cells[20, "B"] as Range).Interior.Color = Color.Blue;
                (xlSht.Cells[20, "C"] as Range).Interior.Color = Color.Blue;
                (xlSht.Cells[20, "B"] as Range).Font.Color = Color.White;
                (xlSht.Cells[20, "D"] as Range).Interior.Color = Color.Blue;

                (xlSht.Cells[20, "F"] as Range).Interior.Color = Color.Blue;
                (xlSht.Cells[20, "F"] as Range).Font.Color = Color.White;
                (xlSht.Cells[20, "G"] as Range).Interior.Color = Color.Blue;
                (xlSht.Cells[20, "H"] as Range).Interior.Color = Color.Blue;
            }
            else if (Six3 == "Желтый" || Six3 == "желтый")
            {
                (xlSht.Cells[20, "B"] as Range).Interior.Color = Color.Yellow;
                (xlSht.Cells[20, "C"] as Range).Interior.Color = Color.Yellow;
                (xlSht.Cells[20, "D"] as Range).Interior.Color = Color.Yellow;

                (xlSht.Cells[20, "F"] as Range).Interior.Color = Color.Yellow;
                (xlSht.Cells[20, "G"] as Range).Interior.Color = Color.Yellow;
                (xlSht.Cells[20, "H"] as Range).Interior.Color = Color.Yellow;
            }
            else if (Six3 == "Коричневый" || Six3 == "коричневый")
            {
                (xlSht.Cells[20, "B"] as Range).Interior.Color = Color.Brown;
                (xlSht.Cells[20, "B"] as Range).Font.Color = Color.White;
                (xlSht.Cells[20, "C"] as Range).Interior.Color = Color.Brown;
                (xlSht.Cells[20, "D"] as Range).Interior.Color = Color.Brown;

                (xlSht.Cells[20, "F"] as Range).Interior.Color = Color.Brown;
                (xlSht.Cells[20, "F"] as Range).Font.Color = Color.White;
                (xlSht.Cells[20, "G"] as Range).Interior.Color = Color.Brown;
                (xlSht.Cells[20, "H"] as Range).Interior.Color = Color.Brown;
            }
            else if (Six3 == "Черный" || Six3 == "черный")
            {
                (xlSht.Cells[20, "B"] as Range).Interior.Color = Color.Black;
                (xlSht.Cells[20, "B"] as Range).Font.Color = Color.White;
                (xlSht.Cells[20, "C"] as Range).Interior.Color = Color.Black;
                (xlSht.Cells[20, "D"] as Range).Interior.Color = Color.Black;

                (xlSht.Cells[20, "F"] as Range).Interior.Color = Color.Black;
                (xlSht.Cells[20, "F"] as Range).Font.Color = Color.White;
                (xlSht.Cells[20, "G"] as Range).Interior.Color = Color.Black;
                (xlSht.Cells[20, "H"] as Range).Interior.Color = Color.Black;
            }
            else if (Six3 == "Красный" || Six3 == "красный")
            {
                (xlSht.Cells[20, "B"] as Range).Interior.Color = Color.Red;
                (xlSht.Cells[20, "B"] as Range).Font.Color = Color.White;
                (xlSht.Cells[20, "C"] as Range).Interior.Color = Color.Red;
                (xlSht.Cells[20, "D"] as Range).Interior.Color = Color.Red;

                (xlSht.Cells[20, "F"] as Range).Interior.Color = Color.Red;
                (xlSht.Cells[20, "F"] as Range).Font.Color = Color.White;
                (xlSht.Cells[20, "G"] as Range).Interior.Color = Color.Red;
                (xlSht.Cells[20, "H"] as Range).Interior.Color = Color.Red;
            }
            else if (Six3 == "Зеленый" || Six3 == "зеленый")
            {
                (xlSht.Cells[20, "B"] as Range).Interior.Color = Color.Green;
                (xlSht.Cells[20, "B"] as Range).Font.Color = Color.White;
                (xlSht.Cells[20, "C"] as Range).Interior.Color = Color.Green;
                (xlSht.Cells[20, "D"] as Range).Interior.Color = Color.Green;

                (xlSht.Cells[20, "F"] as Range).Interior.Color = Color.Green;
                (xlSht.Cells[20, "F"] as Range).Font.Color = Color.White;
                (xlSht.Cells[20, "G"] as Range).Interior.Color = Color.Green;
                (xlSht.Cells[20, "H"] as Range).Interior.Color = Color.Green;
            }
            else if (Six3 == "Серый" || Six3 == "серый")
            {
                (xlSht.Cells[20, "B"] as Range).Interior.Color = Color.Gray;
                (xlSht.Cells[20, "C"] as Range).Interior.Color = Color.Gray;
                (xlSht.Cells[20, "D"] as Range).Interior.Color = Color.Gray;

                (xlSht.Cells[20, "F"] as Range).Interior.Color = Color.Gray;
                (xlSht.Cells[20, "G"] as Range).Interior.Color = Color.Gray;
                (xlSht.Cells[20, "H"] as Range).Interior.Color = Color.Gray;
            }
            else if (Six3 == "Оранжевый" || Six3 == "оранжевый")
            {
                (xlSht.Cells[20, "B"] as Range).Interior.Color = Color.Orange;
                (xlSht.Cells[20, "C"] as Range).Interior.Color = Color.Orange;
                (xlSht.Cells[20, "D"] as Range).Interior.Color = Color.Orange;

                (xlSht.Cells[20, "F"] as Range).Interior.Color = Color.Orange;
                (xlSht.Cells[20, "G"] as Range).Interior.Color = Color.Orange;
                (xlSht.Cells[20, "H"] as Range).Interior.Color = Color.Orange;
            }
            else if (Six3 == "Розовый" || Six3 == "розовый")
            {
                (xlSht.Cells[20, "B"] as Range).Interior.Color = Color.Pink;
                (xlSht.Cells[20, "C"] as Range).Interior.Color = Color.Pink;
                (xlSht.Cells[20, "D"] as Range).Interior.Color = Color.Pink;

                (xlSht.Cells[20, "F"] as Range).Interior.Color = Color.Pink;
                (xlSht.Cells[20, "G"] as Range).Interior.Color = Color.Pink;
                (xlSht.Cells[20, "H"] as Range).Interior.Color = Color.Pink;
            }
            else if (Six3 == "Фиолетовый" || Six3 == "фиолетовый")
            {
                (xlSht.Cells[20, "B"] as Range).Interior.Color = Color.Violet;
                (xlSht.Cells[20, "B"] as Range).Font.Color = Color.White;
                (xlSht.Cells[20, "C"] as Range).Interior.Color = Color.Violet;
                (xlSht.Cells[20, "D"] as Range).Interior.Color = Color.Violet;

                (xlSht.Cells[20, "F"] as Range).Interior.Color = Color.Violet;
                (xlSht.Cells[20, "F"] as Range).Font.Color = Color.White;
                (xlSht.Cells[20, "G"] as Range).Interior.Color = Color.Violet;
                (xlSht.Cells[20, "H"] as Range).Interior.Color = Color.Violet;
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void printPreviewDialog1_Load(object sender, EventArgs e)
        {

        }



        private void printPreviewDialog2_Load(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Text != "")
            {
                textBox2.Enabled = true;
                button2.Enabled = true;
                checkBox1.Enabled = true;
                checkBox2.Enabled = true;
            }
            else if (textBox1.Text == "")
            {
                button2.Enabled = false;
                textBox2.Text = null;
                textBox2.Enabled = false;
                checkBox1.Enabled = false;
                checkBox2.Enabled = false;
            }
        }

        private void textBox2_EnabledChanged(object sender, EventArgs e)
        {
            
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (textBox2.Text != "")
            {
                textBox3.Enabled = true;
                checkBox4.Enabled = true;
                checkBox3.Enabled = true;
            }
            else if (textBox2.Text == "")
            {
                textBox3.Text = null;
                textBox3.Enabled = false;
                checkBox4.Enabled = false;
                checkBox3.Enabled = false;
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked==true)
            {
                checkBox2.Checked = false;
                textBox4.Enabled = true;
                textBox7.Enabled = true;
            }
           
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                checkBox1.Checked = false;
                textBox4.Enabled = false;
                textBox7.Enabled = true;
            }
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked == true)
            {
                checkBox3.Checked = false;
                textBox5.Enabled = true;
                textBox8.Enabled = true;
            }
            
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox6.Checked == true)
            {
                checkBox5.Checked = false;
                textBox6.Enabled = true;
                textBox9.Enabled = true;
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked == true)
            {
                checkBox4.Checked = false;
                textBox5.Enabled = false;
                textBox8.Enabled = true;
            }
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked == true)
            {
                checkBox6.Checked = false;
                textBox6.Enabled = false;
                textBox9.Enabled = true;
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            if (textBox3.Text != "")
            {
            
                checkBox5.Enabled = true;
                checkBox6.Enabled = true;
            }
            else if (textBox3.Text == "")
            {
              
                checkBox6.Enabled = false;
                checkBox5.Enabled = false;
            }
        }
    }
}
