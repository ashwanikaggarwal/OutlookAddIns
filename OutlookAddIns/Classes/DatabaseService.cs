using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
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

        #region CreateBrandNew
        public bool CreateBrandNew(ContactDetails details)
        {
            //create address
            int addID = InsertNewAddress(details.Address, details.Website, details.Instagram, details.Twitter);
            if (addID == 0) return false;

            //create company
            int compID = AddCompany(details, addID);
            if (compID == 0) return false;

            //add note
            if (!AddNote(compID)) return false;

            //create contact
            int contactID = AddContact(details, addID, compID);
            if (contactID == 0) return false;

            //create reg
            if (!AddRegistration(contactID)) return false;

            return true;
        }
        #endregion

        #region UpdateContact
        public bool UpdateContactInfo(int ID, Dictionary<string, string> details)
        {
            //check address, add new one if need, and return id
            int addID = GetAddress(details["Address"], ID);
            if (addID == 0) return false;

            //update contacts
            string SQL = @" PARAMETERS  jt Text, 
                                        mobile Text, 
                                        phone Text, 
                                        deet DateTime, 
                                        addID Long, 
                                        twitter Text;
                            UPDATE TblContacts
                            SET    JobTitle = [jt],
                                   MobileTelNumber = [mobile],
                                   [Direct Tel No] = [phone],
                                   [Opted-In] = true,
                                   [Opted-InDate] = [deet],
                                   AddressID = [addID],
                                   PreviousAddress = [addID],
                                   TblContacts.Twitter = [twitter]
                            WHERE  ContactID = " + ID;

            Access.Dao.QueryDef qDef = db.CreateQueryDef("", SQL);

            qDef.Parameters["jt"].Value = details["Job Title"];
            qDef.Parameters["mobile"].Value = details["Mobile"];
            qDef.Parameters["phone"].Value = details["Phone"];
            qDef.Parameters["deet"].Value = DateTime.Today.ToString("dd/MM/yyyy");
            qDef.Parameters["twitter"].Value = details["Twitter"];
            qDef.Parameters["addID"].Value = addID;

            qDef.Execute();
            qDef.Close();

            return true;
        }
        #endregion

        #region AddCompany
        public int AddCompany(ContactDetails details, int AddressID)
        {
            try
            {
                string PARAMS = "PARAMETERS addID Long, cName Text, cTel Text; ";
                string INSERT = "INSERT INTO TblExhibitors (Company, Tel, CompanyType, AddressID) ";
                string VALUES = "VALUES([cName], [cTel], 'Visitor', [addID]);";
                string SQL = PARAMS + INSERT + VALUES;

                Access.Dao.QueryDef qDef = db.CreateQueryDef("", SQL);
                qDef.Parameters["addID"].Value = AddressID;
                qDef.Parameters["cName"].Value = details.Company;
                qDef.Parameters["cTel"].Value = details.Phone == "" || details.Phone == null ? details.Mobile : details.Phone;

                qDef.Execute();
                qDef.Close();

                return db.OpenRecordset("SELECT @@IDENTITY").Fields[0].Value;
            }
            catch (Exception)
            {

                return 0;
            }

        }
        #endregion

        #region AddNote
        /// <summary>
        /// Adds an initial note to the newly created Company.
        /// </summary>
        /// <param name="CompanyID">ID of the newly created Company.</param>
        /// <returns></returns>
        public bool AddNote(int CompanyID)
        {
            try
            {
                string SQL = $@"INSERT INTO TblNotes (ExhibitorID, UserStamp, Notes)
                                VALUES ({CompanyID}, 'OTLK', 'Ticket Request - Created from Outlook AddIn');";
                db.Execute(SQL);
                return true;
            }
            catch (Exception)
            {

                return false;
            }
        }
        #endregion

        #region AddContact
        /// <summary>
        /// Adds a new Contact provided all the details. This should be the last step when adding a brand new entry.
        /// </summary>
        /// <param name="details">A ContactDetails class with all available contact details. (For registration.) </param>
        /// <param name="AddressID">A valid AddressID previously generated.</param>
        /// <param name="CompanyID">A valid CompanyID previously generated.</param>
        /// <returns></returns>
        public int AddContact(ContactDetails details, int AddressID, int CompanyID)
        {
            try
            {
                string PARAMS = "PARAMETERS fn Text, ln Text, twit Text, jt Text, phone Text, mobile Text, em Text, optDate DateTime; ";
                string INSERT = "INSERT INTO TblContacts (Twitter, [First Name], Surname, JobTitle, MobileTelNumber, Email, " +
                                "Salutation, ExhibitorID, [Direct Tel No], [Opted-In], [Opted-InDate], SameAddress, AddressID) ";
                string VALUES = $@"VALUES ([twit], [fn], [ln], [jt], [mobile], [em], [fn], {CompanyID}, [phone], True, [optDate], True, {AddressID});";
                string SQL = PARAMS + INSERT + VALUES;
                Access.Dao.QueryDef qd = db.CreateQueryDef("", SQL);

                qd.Parameters["fn"].Value = details.FirstName;
                qd.Parameters["ln"].Value = details.LastName;
                qd.Parameters["twit"].Value = details.Twitter;
                qd.Parameters["jt"].Value = details.JobTitle;
                qd.Parameters["phone"].Value = details.Phone;
                qd.Parameters["mobile"].Value = details.Mobile;
                qd.Parameters["em"].Value = details.Email;
                qd.Parameters["optDate"].Value = DateTime.Today.ToString("dd/MM/yyyy");

                qd.Execute();
                qd.Close();

                return db.OpenRecordset("SELECT @@IDENTITY").Fields[0].Value;
            }
            catch (Exception)
            {

                return 0;
            }

        }
        #endregion

        #region AddRegistration

        /// <summary>
        /// Adds a Ticket Registration (without the details) to the database. Marking the Contact as Registered
        /// for the selected campaign.
        /// </summary>
        /// <param name="ID">A valid ContactID. </param>
        /// <returns></returns>
        public bool AddRegistration(int ID)
        {

            //update registrations
            string SQL = $@"INSERT INTO TblRegistrations(ContactID, [Interested?], ExhibitionID) 
                            VALUES({ID} ,true, {Properties.Settings.Default.CurrentVisitorExhibition})";
            db.Execute(SQL);
            return true;

        }
        #endregion

        #region AddAddress
        /// <summary>
        /// Inserts a new address into the Database and retrieves said new address' ID.
        /// </summary>
        /// <param name="toInsert">Address to insert.</param>
        /// <returns></returns>
        public int InsertNewAddress(Address toInsert)
        {
            try
            {
                int countryID = GetCountry(toInsert.Country);
                string SQL = @"PARAMETERS   Address1 Text, 
                                        Address2 Text, 
                                        Address3 Text, 
                                        Town Text, 
                                        County Text, 
                                        Postcode Text, 
                                        Country Short;
                            INSERT INTO TblAddresses (Address1, Address2, 
                            Address3, Town, County, Postcode, Country)
                            VALUES([Address1], [Address2], [Address3], [Town], 
                            [County],[Postcode], [Country]);";
                Access.Dao.QueryDef qInsertDef = db.CreateQueryDef("", SQL);
                //determine parameters
                foreach (PropertyInfo p in typeof(Address).GetProperties())
                {
                    if (p.Name == "Item") continue;
                    if (p.Name == "Country") qInsertDef.Parameters["Country"].Value = countryID;
                    else qInsertDef.Parameters[p.Name].Value = toInsert[p.Name];
                }

                qInsertDef.Execute();
                qInsertDef.Close();



                return db.OpenRecordset("SELECT @@IDENTITY").Fields[0].Value;

            }
            catch (Exception)
            {
                return 0;
            }

        }

        /// <summary>
        /// Inserts a new address into the Database and retrieves said new address' ID. Also inserts the full Social Media bits.
        /// </summary>
        /// <param name="toInsert">Address to insert.</param>
        /// <param name="Website">Website to insert.</param>
        /// <param name="Insta">Insta to insert.</param>
        /// <param name="Twitter">Twitter to insert.</param>
        /// <returns></returns>
        public int InsertNewAddress(Address toInsert, string Website, string Insta, string Twitter)
        {
            try
            {
                int countryID = GetCountry(toInsert.Country);
                string SQL = @"PARAMETERS   Address1 Text, 
                                        Address2 Text, 
                                        Address3 Text, 
                                        Town Text, 
                                        County Text, 
                                        Postcode Text, 
                                        Country Short,
                                        twit Text,
                                        website Text,
                                        insta Text;
                            INSERT INTO TblAddresses (  Address1, 
                                                        Address2, 
                                                        Address3, 
                                                        Town, 
                                                        County, 
                                                        Postcode, 
                                                        Country, 
                                                        Web, 
                                                        Twitter, 
                                                        Instagram) 
                            VALUES( [Address1], 
                                    [Address2], 
                                    [Address3], 
                                    [Town], 
                                    [County],
                                    [Postcode], 
                                    [Country], 
                                    [website], 
                                    [twit], 
                                    [insta]);";

                Access.Dao.QueryDef qInsertDef = db.CreateQueryDef("", SQL);
                //determine parameters
                foreach (PropertyInfo p in typeof(Address).GetProperties())
                {
                    if (p.Name == "Item") continue;
                    if (p.Name == "Country") qInsertDef.Parameters["Country"].Value = countryID;
                    else qInsertDef.Parameters[p.Name].Value = toInsert[p.Name];
                }
                qInsertDef.Parameters["website"].Value = Website;
                qInsertDef.Parameters["twit"].Value = Twitter;
                qInsertDef.Parameters["insta"].Value = Insta;

                qInsertDef.Execute();
                qInsertDef.Close();

                return db.OpenRecordset("SELECT @@IDENTITY").Fields[0].Value;

            }
            catch (Exception)
            {
                return 0;
            }

        }


        #endregion

        #region FindContactReference
        public int FindContactReference(KeyDetails details)
        {
            bool foundSomething = false;

            //all matches?
            int perfect = FindPerfectRegistration(details);
            if ( perfect != 0) return perfect;
            
            if (checkCombinations(CompanyName: details.Company)) { foundSomething = true; return 0; }
            if (checkCombinations(Email: details.Email)) { foundSomething = true; return 0; }
            if (checkCombinations(CompanyName: details.Company, Email: details.Email)) { foundSomething = true; return 0; }
            if (checkCombinations(CompanyName: details.Company, Email: details.Email, FirstName: details.FirstName)) { foundSomething = true; return 0; }
            if (checkCombinations(CompanyName: details.Company, FirstName: details.FirstName)) { foundSomething = true; return 0; }
            if (checkCombinations(CompanyName: details.Company, FirstName: details.FirstName, LastName: details.LastName)) { foundSomething = true; return 0; }
            if (checkCombinations(Email: details.Email, FirstName: details.FirstName)) { foundSomething = true; return 0; }
            if (checkCombinations(Email: details.Email, FirstName: details.FirstName, LastName: details.LastName)) { foundSomething = true; return 0; }
            if (checkCombinations(FirstName: details.FirstName, LastName: details.LastName)) { foundSomething = true; return 0; }
            

            if (foundSomething) return 0;
            return 201;
        }


        private bool checkCombinations(string CompanyName = "", string Email = "", string FirstName = "", string LastName = "")
        {
            int count = 0;
            if (CompanyName != "") count += 1; if (Email != "") count += 2; if (FirstName != "") count += 4; if (LastName != "") count += 8;
            if (count == 0) return false;

            //main string
            string PARAMS = "PARAMETERS email Text, compName Text, fn Text, ln Text; ";
            string SELECT = "SELECT ContactID FROM TblExhibitors INNER JOIN TblContacts ON TblExhibitors.ExhibitorID = TblContacts.ExhibitorID ";
            string WHERE = "WHERE ";
            string AND = "AND ";
            //where clauses
            if (CompanyName != "") WHERE += "TblExhibitors.Company LIKE '*' & [compName] & '*' ";
            if (count == 3 || count == 5 || count == 7 || count == 13) WHERE += AND;

            if (Email != "") WHERE += "TblContacts.Email = [Email] ";
            if (count == 6 || count == 7 || count == 14) WHERE += AND;

            if (FirstName != "") WHERE += "[First Name] = [fn] ";
            if (count == 13 || count == 14 || count == 12) WHERE += AND;

            if (LastName != "") WHERE += "Surname = [ln] ";
            WHERE += ";";

            //final sql
            string SQL = PARAMS + SELECT + WHERE;
            Access.Dao.QueryDef qDef = db.CreateQueryDef("", SQL);

            //params
            qDef.Parameters["compName"].Value = CompanyName;
            qDef.Parameters["email"].Value = Email;
            qDef.Parameters["fn"].Value = FirstName;
            qDef.Parameters["ln"].Value = LastName;

            int result = Useful.getSingleIntFromRecorset(qDef.OpenRecordset());
            if (result != 0) return true;
            return false;

        }

        private bool checkNamesAndPostcode(string FirstName, string LastName, string Postcode)
        {

            string PARAM = "PARAMETERS fn Text, ln Text, pc Text;";
            string SELECT = "SELECT TblContacts.ContactID";
            string FROM = "FROM TblContacts LEFT JOIN tblAddresses ON TblContacts.AddressID = tblAddresses.AddressID ";
            string WHERE = "WHERE [First Name] = [fn] AND Surname = [ln] AND Postcode = [pc];";
            string SQL = PARAM + SELECT + FROM + WHERE;
            Access.Dao.QueryDef qDef = db.CreateQueryDef("", SQL);

            //params
            qDef.Parameters["fn"].Value = FirstName;
            qDef.Parameters["ln"].Value = LastName;
            qDef.Parameters["pc"].Value = Postcode;

            int result = Useful.getSingleIntFromRecorset(qDef.OpenRecordset());
            if (result != 0) return true;
            return false;

        }

        #endregion

        #region FindPerfectRegistration
        public int FindPerfectRegistration(KeyDetails details)
        {

            int id = 0;
            int count = 0;
            
            //looks messier but should be sanitized
            string SQL = @"PARAMETERS   fn Text, 
                                        ln Text, 
                                        email Text, 
                                        company Text;
                            SELECT contactid
                            FROM   tblexhibitors
                                   INNER JOIN tblcontacts
                                           ON tblexhibitors.exhibitorid = tblcontacts.exhibitorid 
                            WHERE  tblcontacts.[First Name] = [fn]
                                   AND tblcontacts.surname = [ln]
                                   AND tblcontacts.email = [email]
                                   AND tblexhibitors.company LIKE '*' & [company] & '*';";
            Access.Dao.QueryDef qDef = db.CreateQueryDef("",SQL);
            
            qDef.Parameters["fn"].Value = details.FirstName;
            qDef.Parameters["ln"].Value = details.LastName;
            qDef.Parameters["email"].Value = details.Email; ;
            qDef.Parameters["company"].Value = details.Company;


            Access.Dao.Recordset rs = qDef.OpenRecordset();

            if (rs.EOF && rs.BOF) return 0;
            rs.MoveFirst();
            
            while(rs.EOF == false)
            {
                id = rs.Fields["ContactID"].Value;
                count++;
                rs.MoveNext();
            }

            if (count > 1) return 0;
            return id;

        }
        #endregion

        #region Address Bit
        /// <summary>
        /// This method converts the Address argument into an Address Object and tries
        /// to find out if it's similar to the one existant in the Contact's record.
        /// If it finds sufficient similarities it returns the existing address (thus no changes made).
        /// If it finds the addresses are different it creates a new address and retrieves its ID.
        /// </summary>
        /// <param name="address">The Address to check for.</param>
        /// <param name="ID">The ContactID to find the address to compare.</param>
        /// <returns></returns>
        public int GetAddress(string address, int ID)
        {
            //we check the EXISTING address
            Address toSend = SerializetStringAddress(address);
            Address toCompare = GetCurrentAddress(GetAddressID(ID));

            // EXACT same address no need to change
            if (CompareAddresses(toSend, toCompare))
            {
                return GetAddressID(ID);
            }
           

            //we proceed and insert a new address!
            int toReturn = InsertNewAddress(toSend);
            UpdateSameAddress(ID);

            if (toReturn == 0) return 0;
            return toReturn;

        }

        /// <summary>
        /// This method transforms an Address string (in the web format) into an Address Class.
        /// </summary>
        /// <param name="address">The Address to transform.</param>
        /// <returns></returns>
        public Address SerializetStringAddress(string address)
        {
            Address final = new Address();
            string[] arr = address.Split(',');
            Array.Reverse(arr);
            for(int i = 0; i<arr.Length; i++) { arr[i] = arr[i].Trim(); }

            final.Country = GetWebCountry(arr[0]);
            final.Postcode = arr[1];
            final.Address1 = arr[arr.Length - 1];
            for (int i = 2; i < arr.Length - 1; i++)
            {
                if (final.County == null) { final.County = arr[i]; continue; }
                if (final.Town == null) { final.Town = arr[i]; continue; }
                if (final.Address2 == null) { final.Address2 = arr[i]; continue; }
                if (final.Address3 == null) { final.Address3 = arr[i]; continue; }
            }

            return final;
        }

        /// <summary>
        /// Returns an Address ID given a Contact or ExhibitorID. A boolean is used to flag either way.
        /// </summary>
        /// <param name="ID">A ContactID or ExhibitorID</param>
        /// <param name="TrueIfContact">Optional. True for ContactID, False for ExhibitorID. Defaults to true.</param>
        /// <returns></returns>
        private int GetAddressID(int ID, bool TrueIfContact = true)
        {
            string SQL;
            if (TrueIfContact) { SQL = "SELECT AddressID FROM TblContacts WHERE ContactID = " + ID; }
            else { SQL = "SELECT AddressID FROM TblExhibitors WHERE ExhibitorID = " + ID; }

            Access.Dao.Recordset rs = db.OpenRecordset(SQL);

            if (rs.EOF && rs.BOF) return 0;
            rs.MoveFirst();

            return rs.Fields["AddressID"].Value;

        }

        /// <summary>
        /// Returns a full Address as an Address Class. Used Internally.
        /// </summary>
        /// <param name="AddressID">A valid AddressID.</param>
        /// <returns></returns>
        private Address GetCurrentAddress(int AddressID)
        {
            Address final = new Address();
            string SQL = "SELECT Address1, Address2, Address3, Town, County, Postcode, Country FROM TblAddresses WHERE AddressID = " + AddressID;
            Access.Dao.Recordset rs = db.OpenRecordset(SQL);

            if (rs.EOF && rs.BOF) return new Address();

            rs.MoveFirst();

            foreach (PropertyInfo p in typeof(Address).GetProperties())
            {
                if (p.Name == "Item") continue;
                if (p.Name == "Country") { final[p.Name] = Convert.IsDBNull(rs.Fields[p.Name].Value) ? "" : GetCountry(rs.Fields[p.Name].Value); }
                else { final[p.Name] = Convert.IsDBNull(rs.Fields[p.Name].Value) ? "" : rs.Fields[p.Name].Value; }
                
            }

            return final;

        }

        /// <summary>
        /// Gets the country name.
        /// </summary>
        /// <param name="countryID">A valid CountryCodeId</param>
        /// <returns></returns>
        public string GetCountry(int countryID)
        {
            string SQL = "SELECT Country FROM TblCountries WHERE CountryCodeId = " + countryID;
            Access.Dao.Recordset rs = db.OpenRecordset(SQL);

            if (rs.EOF && rs.BOF) return "error";

            return rs.Fields["Country"].Value;
        }
        /// <summary>
        /// Gets the country code.
        /// </summary>
        /// <param name="countryName">A matching Country name.</param>
        /// <returns></returns>
        public int GetCountry(string countryName)
        {
            string SQL = "SELECT CountryCodeId FROM TblCountries WHERE Country = '" + countryName + "'";
            Access.Dao.Recordset rs = db.OpenRecordset(SQL);

            if (rs.EOF && rs.BOF) return 0;

            return rs.Fields["CountryCodeId"].Value;
        }

        /// <summary>
        /// Gets the Country name given the country web code found in the address given by the website's form.
        /// </summary>
        /// <param name="CountryCode"></param>
        /// <returns></returns>
        private string GetWebCountry(string CountryCode)
        {
            
            string SQL = "SELECT country FROM TblWebCountries WHERE value = '" + CountryCode + "';";
            Access.Dao.Recordset rs = db.OpenRecordset(SQL);
            return rs.Fields["country"].Value;
        }

        /// <summary>
        /// Updates the tickbox "SameAddress" that's used in the Front End to replace the address with the Company's address.
        /// </summary>
        /// <param name="ID">A valid ContactID</param>
        private void UpdateSameAddress(int ID)
        {
            string SQL = "UPDATE TblContacts SET SameAddress = False WHERE ContactID = " + ID;
            db.Execute(SQL);
        }

        /// <summary>
        /// Compares two given addresses. It search for a match of Postcode as a baseline and then
        /// tries to gather as many matches possible. If it passes the threshold it returns as equal.
        /// </summary>
        /// <param name="address1">First Address to compare.</param>
        /// <param name="address2">Second Address to compare.</param>
        /// <returns></returns>
        private bool CompareAddresses(Address address1, Address address2)
        {
            //for comparison, the Postcode needs to match, then at least 5/7
            if (address1.Postcode == address2.Postcode)
            {
                int match = 0;
                List<string> arr = new List<string>();

                //first we put the address1 into an array/list
                foreach (PropertyInfo p in typeof(Address).GetProperties())
                {
                    if (p.Name == "Item") continue;

                    string toAdd;
                    toAdd = address1[p.Name] == "" | address1[p.Name] == null ? 
                                "" : string.Join("", address1[p.Name].Split(default(string[]), StringSplitOptions.RemoveEmptyEntries)).ToUpper();
                    
                    arr.Add(toAdd);
                }

                //then we compare them
                foreach (var item in address2)
                {

                    string toCheck;
                    toCheck = item == "" | item == null ?
                                "&%NOMATCH%&" : string.Join("", item.Split(default(string[]), StringSplitOptions.RemoveEmptyEntries)).ToUpper();

                    if (arr.Contains(toCheck)) match++;
                }

                //if match is 4 or more then we say its same
                if (match >= 4) return true;
                return false;

            }
            else { return false; }
           
        }

        #endregion
    }
}
