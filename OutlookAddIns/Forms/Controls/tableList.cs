using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutlookAddIns.Forms.Controls
{
    public partial class tableList : UserControl
    {
        public string TableName
        {
            private get { return TableName; }
            set
            {
                txtTableName.Text = TableName;
            }
        }

        public string[] Fields
        {
            get { return Fields; }
            set { lstFormFields.DataSource = Fields; }
        }

        public string[] DbFields
        {
            get { return DbFields; }
            set { lstDbFields.DataSource = DbFields; }
        }

        public tableList()
        {
            InitializeComponent();
        }
    }
}
