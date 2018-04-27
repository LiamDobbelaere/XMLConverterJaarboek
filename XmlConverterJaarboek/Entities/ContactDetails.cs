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
            if (StreetName != "") details.Add(StreetName + (StreetNumber != "" ? Characters.FIXED_SPACE + StreetNumber : "") + (Box != "" ? Characters.FIXED_SPACE + Box : ""));
            if (PostalCode != "") details.Add(PostalCode + (Town != "" ? Characters.FIXED_SPACE + Town : ""));
            if (Telephone != "" && Telephone.Equals(Fax))
            {
                if (Telephone != "")
                {
                    details.Add("T/F" + Characters.FIXED_SPACE + Telephone.Replace("-", Characters.NOBREAK_HYPHEN));
                }
            }
            else
            {
                if (Telephone != "")
                {
                    details.Add("T" + Characters.FIXED_SPACE + Telephone.Replace("-", Characters.NOBREAK_HYPHEN));
                }
                if (Fax != "")
                {
                    details.Add("F" + Characters.FIXED_SPACE + Fax.Replace("-", Characters.NOBREAK_HYPHEN));
                }
            }
            if (Cellphone != "")
            {
                details.Add("G" + Characters.FIXED_SPACE + Cellphone.Replace("-", Characters.NOBREAK_HYPHEN));
            }
            if (Email != "")
            {
                details.Add(Email);
            }

            return string.Join(Characters.FIXED_SPACE + Characters.NOBREAK_HYPHEN + Characters.FIXED_SPACE, details.ToArray());
        }
    }
}
