using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XmlConverterJaarboek.Entities
{
    class ContactDetails
    {
        public string Institution { get; set; }
        public string StreetName { get; set; }
        public string StreetNumber { get; set; }
        public string Box { get; set; }
        public string PostalCode { get; set; }
        public string Town { get; set; }
        public string Telephone { get; set; }
        public string Fax { get; set; }
        public string Cellphone { get; set; }
        public string Email { get; set; }

        public string GetFormattedDetails()
        {
            List<string> details = new List<string>();

            if (Institution != "") details.Add(Institution);
            if (StreetName != "") details.Add(StreetName + " " + StreetNumber);
            if (PostalCode != "") details.Add(PostalCode + " " + Town);
            if (Telephone != "")
            {
                details.Add("T " + Telephone);
            }
            if (Fax != "")
            {
                details.Add("F " + Fax);
            }
            if (Cellphone != "")
            {
                details.Add("G " + Cellphone);
            }

            return string.Join(" - ", details.ToArray());
        }
    }
}
