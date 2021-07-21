using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Access = Microsoft.Office.Interop.Access;

namespace OutlookAddIns.Classes
{
    public static class Useful
    {

        public static string getSingleStrFromRecorset( Access.Dao.Recordset recordset)
        {
            if (recordset.EOF && recordset.BOF) return "NORECORDS";
            recordset.MoveFirst();

            return recordset.Fields[0].Value;
        }

        public static int getSingleIntFromRecorset(Access.Dao.Recordset recordset)
        {
            if (recordset.EOF && recordset.BOF) return 0;
            recordset.MoveFirst();

            return recordset.Fields[0].Value;
        }

    }
}
