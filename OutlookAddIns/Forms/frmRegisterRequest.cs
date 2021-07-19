using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OutlookAddIns.Classes;

namespace OutlookAddIns.Forms
{
    public partial class frmRegisterRequest : Form
    {
        private Dictionary<string, object> details;
        public frmRegisterRequest(Dictionary<string, object> Details)
        {
            InitializeComponent();
            details = Details;
            string rtfText;
            rtfText = @"{\rtf1\ansi";
            foreach (KeyValuePair<string, object> detail in Details)
            {

                rtfText += @"\b " + detail.Key + @"\b0   " + detail.Value + @" \line";
            }
            rtfText += "}";
            txtValuesRich.Rtf = rtfText;
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            List<int> test = new DatabaseService().EmailLookUp(details["Email"].ToString());
            MessageBox.Show("" + test[0]);
        }
    }
}
