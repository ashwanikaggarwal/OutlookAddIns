using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;

namespace OutlookAddIns.Classes
{
    class ContactDetails : IEnumerable<object>
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string JobTitle { get; set; }
        public string Company { get; set; }
        public Address Address { get; set; }
        public int AddressID { get; set; }
        public string Phone { get; set; }
        public string Mobile { get; set; }
        public string Email { get; set; }
        public string Twitter { get; set; }
        public string Website { get; set; }
        public string Instagram { get; set; }

        public void SetFromDictionary(Dictionary<string, string> details)
        {
            //First Name;Last Name;Job Title;Address;Phone;Mobile;Email;Website;Twitter;Instagram;Company on Badge
            FirstName = details["First Name"];
            LastName = details["Last Name"];
            JobTitle = details["Job Title"];
            Company = details["Company on Badge"];
            DatabaseService db = new DatabaseService();
            Address = db.SerializetStringAddress(details["Address"]);
            Phone = details["Phone"];
            Mobile = details["Mobile"];
            Email = details["Email"];
            Twitter = details["Twitter"];
            Website = details["Website"];
            Instagram = details["Instagram"];
        }

        #region [] Getter/Setter
        public string this[string propertyName]
        {
            set
            {
                switch (propertyName)
                {
                    case "FirstName":
                        FirstName = value;
                        break;
                    case "LastName":
                        LastName = value;
                        break;
                    case "JobTitle":
                        JobTitle = value;
                        break;
                    case "Company":
                        Company = value;
                        break;
                    case "Phone":
                        Phone = value;
                        break;
                    case "Mobile":
                        Mobile = value;
                        break;
                    case "Email":
                        Email = value;
                        break;
                    case "Twitter":
                        Twitter = value;
                        break;
                }
            }
            get
            {
                switch (propertyName)
                {
                    case "FirstName":
                        return FirstName;

                    case "LastName":
                        return LastName;

                    case "JobTitle":
                        return JobTitle;

                    case "Company":
                        return Company;

                    case "Phone":
                        return Phone;

                    case "Mobile":
                        return Mobile;

                    case "Email":
                        return Email;

                    case "Twitter":
                        return Twitter;

                    default:
                        return "ERROR";
                }
            }
        }
        #endregion

        #region Enumerator
        public IEnumerator<object> GetEnumerator()
        {
            yield return FirstName;
            yield return LastName;
            yield return JobTitle;
            yield return Company;
            yield return Address;
            yield return Phone;
            yield return Mobile;
            yield return Email;
            yield return Twitter;


        }
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
        #endregion

    }
}
