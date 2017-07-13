using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Module_Report;

namespace ReportDemo
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create new tab to place content on
            TabPage Page1 = new TabPage("Reporting");
            Page1.Height = 400;
            Page1.Width = 265;
            tabControl1.TabPages.Add(Page1);

            //create a panel to set equal to existing panel with numerous controls & event handlers on
            Report Report_Panel = new Report();
            Panel Panel = Report_Panel.panel1;

            //Put the panel control on the dynamic generated tab page control
            Page1.Controls.Add(Panel);
        }
    }
}
