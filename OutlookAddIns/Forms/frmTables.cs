using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OutlookAddIns.Forms.Controls;

namespace OutlookAddIns.Forms
{
    public partial class frmTables : Form
    {
        public frmTables()
        {
            InitializeComponent();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            tableList list = new tableList();
            list.Show();
            list.Location = new System.Drawing.Point(50, 50);
            
        }
    }
}
