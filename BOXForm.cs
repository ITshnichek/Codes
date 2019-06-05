using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PackingTool
{
    public partial class BOXForm : Form
    {
        public BOXForm()
        {
            InitializeComponent();
            Info.S = "";
            Info.P = "";
            Info.Q = "";
            Info.SS = "";

        }
        
        private void BOXForm_Load(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = Info.Name;
            timer1.Enabled = true;
            timer1.Interval = 1000;

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            DateTime ci = DateTime.Now;
            toolStripStatusLabel2.Text = ci.ToLongTimeString();
            tbScanGalia.Focus();
        }

        private void BOXForm_FormClosed(object sender, FormClosedEventArgs e)
        {
                  

        }

        private void tbScanGalia_KeyDown(object sender, KeyEventArgs e)
        {
            DBConnect con = new DBConnect();
            DataTable box = new DataTable();
            DataTable box1 = new DataTable();
            if (e.KeyCode == Keys.Enter)
            {
                if(tbScanGalia.Text.Contains("P"))
                {                    
                        Info.P = "1";
                        Info.CodeProduct = tbScanGalia.Text.TrimStart('P');
                        CodeProduct();
                        tbScanGalia.Clear();
                        OpenPackingForm();
                    }
                    
                if (tbScanGalia.Text.Contains("S") && tbScanGalia.Text.Length == 10)
                {
                box = con.Select("galianumber, date", "packingbox", "galianumber = '" + tbScanGalia.Text.TrimStart('S') + "' AND status = True");
                box1 = con.Select("galianumber, qtpacked, qttotal, date, psanumber, leoninumber", "packingbox", "galianumber = '" + tbScanGalia.Text.TrimStart('S') + "' AND status = False");
                    if (box.Rows.Count == 0) //проверка на повтор коробки
                    {
                        if (box1.Rows.Count == 0) //проверка на не закрытую коробку
                        {
                            Info.S = "1";
                            Info.GaliaNumber = tbScanGalia.Text.TrimStart('S');
                            Eti();
                            tbScanGalia.Clear();
                            OpenPackingForm();
                        }
                        else
                        {
                            Info.QTPacking = Convert.ToInt32(box1.Rows[0][2].ToString());
                            Info.CounterHRNS = Convert.ToInt32(box1.Rows[0][1].ToString());
                            Info.CodeProduct = box1.Rows[0][4].ToString();
                            Info.REF30S = box1.Rows[0][5].ToString();
                            Info.GaliaNumber = box1.Rows[0][0].ToString();
                            Info.S = "1";
                            Info.P = "1";
                            Info.SS = "1";
                            Info.Q = "1";
                            OpenPackingForContinue();
                        }
                    }
                    else
                    {
                        Info.ParamError = "Повторное использование ! Коробка с номером "+ box.Rows[0][0].ToString() + " уже упакована "+ box.Rows[0][1].ToString();
                        errorBOX errb = new errorBOX();
                        this.Hide();
                        errb.Show();
                        //MessageBox.Show("ERROR!!!!!!");
                        
                    }
                    
                }
                if (tbScanGalia.Text.Contains("Q"))
                {
                    Info.Q = "1";
                    Info.QTPacking = Convert.ToInt32(tbScanGalia.Text.TrimStart('Q'));
                    Qty();
                    tbScanGalia.Clear();
                    OpenPackingForm();
                }
                if (tbScanGalia.Text.Contains("30S"))
                {
                    Info.SS = "1";
                    Info.REF30S = tbScanGalia.Text.TrimStart('3').TrimStart('0').TrimStart('S');
                    Ref();
                    tbScanGalia.Clear();
                    OpenPackingForm();
                }
                tbScanGalia.Clear();
            }
        }
        private void CodeProduct()
        {
            Pen mypen = new Pen(Color.Green, 4);
            Graphics g = Graphics.FromHwnd(pictureBox1.Handle);
            g.DrawRectangle(mypen, 23, 117, 305, 37);
        }
        private void Qty()
        {
            Pen mypen = new Pen(Color.Green, 3);
            Graphics g = Graphics.FromHwnd(pictureBox1.Handle);
            g.DrawRectangle(mypen, 23, 195, 119,40 );
        }
        private void Eti()
        {
            Pen mypen = new Pen(Color.Green, 3);
            Graphics g = Graphics.FromHwnd(pictureBox1.Handle);
            g.DrawRectangle(mypen, 23, 316, 281, 41);
        }
        private void Ref()
        {
            Pen mypen = new Pen(Color.Green, 3);
            Graphics g = Graphics.FromHwnd(pictureBox1.Handle);
            g.DrawRectangle(mypen, 355, 227, 280, 40);
        }
        private void OpenPackingForm()
        {
            
            if (Info.S == "1" && Info.P == "1" && Info.SS == "1" && Info.Q == "1" )
            {
                DBConnect con = new DBConnect();
                con.Insert("packingbox", "galianumber, qttotal, psanumber, leoninumber",Info.GaliaNumber+", "+Info.QTPacking+", '"+ Info.CodeProduct+"', '"+ Info.REF30S+"'");
                this.Close();
                Properties.Settings.Default.project = "p32s";
                Properties.Settings.Default.semeistvo = "main";
                Properties.Settings.Default.referencia = "P" + Info.CodeProduct;
                ActivePoint pack = new ActivePoint();
                pack.Show();
            }
        }
       
        private void OpenPackingForContinue()
        {

            if (Info.S == "1" && Info.P == "1" && Info.SS == "1" && Info.Q == "1")
            {
                this.Close();
  
                ActivePoint pack = new ActivePoint();
                pack.Show();
            }
        }

        private void btChangeOperator_Click(object sender, EventArgs e)
        {
            authForm authf = new authForm();
            if (Application.OpenForms["authf"] == null)
            {

                authf.Show();
            }
            else
            {
                Application.OpenForms["authf"].Focus();
            }



            //authForm authf = new authForm();
            //authf.Show();
            this.Close();
            
            
        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void pictureBox1_MouseClick(object sender, MouseEventArgs e)
        {
            int x = e.Location.X;
            int y = e.Location.Y;
        }
    }
}
