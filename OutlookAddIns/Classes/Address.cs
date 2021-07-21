using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookAddIns.Classes
{
    class Address : IEnumerable<string>
    {
        public string Country { get; set; }
        public string Postcode { get; set; }
        public string County { get; set; }
        public string Town { get; set; }
        public string Address3 { get; set; }
        public string Address2 { get; set; }
        public string Address1 { get; set; }

        public string this[string propertyName]
        {
            set
            {
                switch (propertyName)
                {
                    case "Country":
                        Country = value;
                        break;
                    case "Postcode":
                        Postcode = value;
                        break;
                    case "County":
                        County = value;
                        break;
                    case "Town":
                        Town = value;
                        break;
                    case "Address3":
                        Address3 = value;
                        break;
                    case "Address2":
                        Address2 = value;
                        break;
                    case "Address1":
                        Address1 = value;
                        break;
                }
            }
            get
            {
                switch (propertyName)
                {
                    case "Country":
                        return Country;

                    case "Postcode":
                        return Postcode;
       
                    case "County":
                        return County;
       
                    case "Town":
                        return Town;
       
                    case "Address3":
                       return Address3;
       
                    case "Address2":
                        return Address2;
       
                    case "Address1":
                        return Address1;

                    default:
                        return "ERROR";
                }
            }
        }

        public IEnumerator<string> GetEnumerator()
        {
            yield return Country;
            yield return Postcode;
            yield return County;
            yield return Town;
            yield return Address3;
            yield return Address2;
            yield return Address1;
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
