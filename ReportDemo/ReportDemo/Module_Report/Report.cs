using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Module_Report
{
    public partial class Report : Form
    {
        ControlEvents_Report EventsClass_Report = new ControlEvents_Report();
        public Report()
        {
            InitializeComponent();
            EventsClass_Report.bindBrowser(lstBrowser);
            EventsClass_Report.bindMapping(lstMapping);
            EventsClass_Report.bindLoad(btnLoad);
            EventsClass_Report.bindGenerate(btnGenerate);
        }
    }
}
