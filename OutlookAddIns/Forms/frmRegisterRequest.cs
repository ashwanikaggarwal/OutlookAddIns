using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using OutlookAddIns.Classes;

namespace OutlookAddIns.Forms
{
    public partial class frmRegisterRequest : Form
    {
        private Dictionary<string, object> details;
        public frmRegisterRequest(Dictionary<string, object> Details)
        {
            InitializeComponent();

            //json test
            Table table = JsonConvert.DeserializeObject<Table>(Properties.Settings.Default.mainJSON);
            textBox1.Text = table.TableName;
            List<string> arr1 = new List<string>();
            List<string> arr2 = new List<string>();
            foreach (var item in table.TableFields)
            {
                arr1.Add(item.FormField);
                arr2.Add(item.DbField);
            }
            listBox1.DataSource = arr1;
            listBox2.DataSource = arr2;
                       

            details = Details;
            //List<bool> checklist = new List<bool>();
            //List<int> test = new DatabaseService().EmailLookUp(details["Email"].ToString());
            //string rtfText;
            //rtfText = @"{\rtf1\ansi";
            //foreach (KeyValuePair<string, object> detail in Details)
            //{

            //    rtfText += @"\b " + detail.Key + @"\b0   " + detail.Value + @" \line";
            //}
            //rtfText += "}";
            //txtValuesRich.Rtf = rtfText;
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            
        }
    }
}
