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
    public partial class packing : Form
    {
        int i;
        public packing()
        {
            
            InitializeComponent();
            labelQTPacking.Text = Info.QTPacking.ToString();
            labelCounterHRNS.Text = Info.CounterHRNS.ToString();
            i = Properties.Settings.Default.timer;
            labelTimer.Text = i.ToString();
            timer1.Enabled = true;
            timer1.Interval = 1000;

        }

        private void timer1_Tick(object sender, EventArgs e)
        {

            labelTimer.Text = (--i).ToString();
            if (i < 0)
            {
                timer1.Stop();
                ActivePoint hform = new ActivePoint();
                hform.Show();
                this.Close();
            }
        }

        private void packing_FormClosed(object sender, FormClosedEventArgs e)
        {
            //BOXForm dbs = new BOXForm();
            //dbs.Show();
        }

        private void labelQTPacking_Click(object sender, EventArgs e)
        {

        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
