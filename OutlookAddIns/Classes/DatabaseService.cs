using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Access = Microsoft.Office.Interop.Access;

namespace OutlookAddIns.Classes
{
    class DatabaseService
    {
        private readonly Properties.Settings settings = Properties.Settings.Default;
        private Access.Dao.Database db;

        public DatabaseService()
        {
            db = new Access.Dao.DBEngine().OpenDatabase(settings.Database);
        }
        

        #region EmailLookUp
        /// <summary>
        /// Tries to find a ContactID given a provided Email. 
        /// <para>Returns a ContactID or 0 if not found.</para>
        /// </summary>
        /// <param name="email">The email to search by.</param>
        /// <returns>ContactID or 0 if not found.</returns>
        public List<int> EmailLookUp(string email)
        {
            List<int> contactIDs = new List<int>();

            string SQL = "SELECT ContactID FROM TblContacts WHERE Email ='" + email + "';";
            Access.Dao.Recordset rs = db.OpenRecordset(SQL);

            if (rs.EOF && rs.BOF) return new List<int>() { 0 };
            rs.MoveFirst();

            while (rs.EOF == false)
            {
                contactIDs.Add(rs.Fields["ContactID"].Value);
                rs.MoveNext();
            }

            return contactIDs;
        }
        #endregion

    }
}
