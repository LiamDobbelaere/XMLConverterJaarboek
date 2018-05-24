using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XmlConverterJaarboek.Entities
{
    class Doctor
    {
        public Doctor()
        {
            ContactDetails = new List<ContactDetails>();
        }

        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Language { get; set; }
        public string INAMI { get; set; }
        public string InExtenso { get; set; }
        public string Competence1 { get; set; }
        public string Competence2 { get; set; }
        public List<ContactDetails> ContactDetails { get; set; }

        public bool IsSameDoctor(string firstName, string lastName, string inami)
        {
            return firstName.Trim().ToLower().Equals(FirstName.Trim().ToLower()) 
                && lastName.Trim().ToLower().Equals(lastName.Trim().ToLower()) 
                && INAMI.Trim().Equals(inami.Trim());
        }

        public string GetFormattedDetails()
        {
            if (Competence1 != "")
            {
                List<string> competences = new List<string>();
                competences.Add(Competence1);
                if (Competence2 != "") competences.Add(Competence2);

                string competenceString = string.Join(Characters.NOBREAK_HYPHEN, competences.ToArray());

                return string.Format("{0}" + Characters.FIXED_SPACE + "{1}\t{2}" + Characters.FIXED_SPACE + "•" 
                    + Characters.FIXED_SPACE + "{3}" + Characters.FIXED_SPACE + "•" + Characters.FIXED_SPACE + "{4}", 
                    LastName, FirstName, competenceString, Language, INAMI);
            }
            else
            {
                return string.Format("{0}" + Characters.FIXED_SPACE + "{1}\t{2}" 
                    + Characters.FIXED_SPACE + "•" + Characters.FIXED_SPACE + "{3}", LastName, FirstName, Language, INAMI);
            }
        }
    }
}
