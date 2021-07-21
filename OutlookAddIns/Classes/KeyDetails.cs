using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookAddIns.Classes
{
    class KeyDetails
    {
        public string Email { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Company { get; set; }

        public string this[string propertyName]
        {
            set
            {
                switch (propertyName)
                {
                    case "Email":
                        Email = value;
                        break;
                    case "FirstName":
                        FirstName = value;
                        break;
                    case "LastName":
                        LastName = value;
                        break;
                    case "Company":
                        Company = value;
                        break;
                }
            }
            get
            {
                switch (propertyName)
                {
                    case "Email":
                        return Email;

                    case "FirstName":
                        return FirstName;

                    case "LastName":
                        return LastName;

                    case "Company":
                        return Company;

                    default:
                        return "ERROR";
                }
            }
        }
    }


}
