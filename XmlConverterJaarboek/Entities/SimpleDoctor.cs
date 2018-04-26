using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XmlConverterJaarboek.Entities
{
    class SimpleDoctor
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string PostalCode { get; set; }
        public string Town { get; set; }
        public string CSouMS { get; set; }
        public string INAMI { get; set; }

        public bool IsSameDoctor(string firstName, string lastName, string inami)
        {
            return firstName.Equals(FirstName) && lastName.Equals(lastName) && INAMI.Equals(inami);
        }
    }
}
