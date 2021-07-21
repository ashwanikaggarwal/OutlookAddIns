using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookAddIns.Classes
{
    class Table
    {
        public string TableName { get; set; }
        public TableContents[] TableFields { get; set; }
    }
}
