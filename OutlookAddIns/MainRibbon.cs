using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using OutlookAddIns.Forms;

namespace OutlookAddIns
{
    public partial class MainRibbon
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnSetPath_Click(object sender, RibbonControlEventArgs e)
        {
            frmSetPath form = new frmSetPath();
            form.Show();
        }
    }
}
