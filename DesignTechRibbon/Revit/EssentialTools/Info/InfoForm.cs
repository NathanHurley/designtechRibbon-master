﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DesignTechRibbonPaid.Revit.EssentialTools.Info
{
    public partial class InfoForm : Form
    {
        public InfoForm()
        {
            InitializeComponent();
        }

        private void designtechLogo_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://designtech.io/");
        }
    }
}
