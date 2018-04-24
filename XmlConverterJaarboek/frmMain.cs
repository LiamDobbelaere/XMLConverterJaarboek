using System;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Xml;
using System.Collections.Generic;
using XmlConverterJaarboek.Entities;

namespace XmlConverterJaarboek
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        private void FormLoad(object sender, EventArgs e)
        {
            using (OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db.mdb"))
            {
                conn.Open();

                StartConversion(conn);

                conn.Close();
            }

            Close();
        }

        private void StartConversion(OleDbConnection conn)
        {
            var writer = XmlWriter.Create("test.xml");
            writer.WriteStartElement("Root");

            writer.WriteElementString("_09_Einde", "");
            ConvertDoctorsForAffiliation(conn, "AP", "MS", writer);
            ConvertDoctorsForAffiliation(conn, "AP", "CS", writer);
            writer.WriteElementString("_09_Einde", "");

            writer.WriteEndElement();
            writer.Flush();
        }

        private void ConvertDoctorsForAffiliation(OleDbConnection conn, string affiliation, string csms, XmlWriter writer)
        {
            bool titleDone = false;
            foreach (Doctor doctor in GetDoctorsForAffiliation(conn, affiliation, csms))
            {
                if (csms == "CS" && !titleDone)
                {
                    writer.WriteElementString("_04_Titel", "CANDIDATS SPECIALISTES•KANDIDATEN-SPECIALISTEN");
                    titleDone = true;
                }

                writer.WriteElementString("_02_Naam", doctor.GetFormattedDetails());

                bool first = true;
                foreach (ContactDetails contactDetails in doctor.ContactDetails)
                {
                    var elementName = first ? "_03_Gegevens1" : "_03_Gegevens2";

                    writer.WriteElementString(elementName, contactDetails.GetFormattedDetails());

                    first = false;
                }
            }
        }

        private List<Doctor> GetDoctorsForAffiliation(OleDbConnection conn, string affiliation, string csms)
        {
            var command = conn.CreateCommand();
            command.CommandText = Queries.DOCTORS_FOR_AFFILIATION;
            command.Parameters.AddRange(new OleDbParameter[]
            {
               new OleDbParameter("@affiliation", affiliation),
               new OleDbParameter("@csms", csms)
            });

            List<Doctor> doctorList = new List<Doctor>();
            var reader = command.ExecuteReader();
            Doctor newDoctor = null;
            while (reader.Read())
            {
                if (newDoctor != null 
                    && newDoctor.IsSameDoctor(
                        reader["PRENOM"].ToString(), 
                        reader["NOM"].ToString(), 
                        reader["NoINAMI"].ToString()))
                {
                    newDoctor.ContactDetails.Add(CreateContactDetails(reader));
                }
                else
                {
                    newDoctor = new Doctor
                    {
                        FirstName = reader["PRENOM"].ToString(),
                        LastName = reader["NOM"].ToString(),
                        INAMI = reader["NoINAMI"].ToString(),
                        Language = reader["Langue"].ToString(),
                        Affiliation = reader["Affiliation"].ToString(),
                        Competence1 = reader["Compétence1"].ToString(),
                        Competence2 = reader["Compétence2"].ToString()
                    };

                    newDoctor.ContactDetails.Add(CreateContactDetails(reader));

                    doctorList.Add(newDoctor);
                }
            }
            reader.Close();

            return doctorList;
        }

        private ContactDetails CreateContactDetails(OleDbDataReader reader)
        {
            return new ContactDetails
            {
                Institution = reader["Institution"].ToString(),
                StreetName = reader["Rue"].ToString(),
                StreetNumber = reader["Nr"].ToString(),
                Box = reader["Bte"].ToString(),
                PostalCode = reader["Poste"].ToString(),
                Town = reader["Commune"].ToString(),
                Telephone = reader["Tel"].ToString(),
                Fax = reader["Fax"].ToString(),
                Cellphone = reader["GSM"].ToString(),
                Email = reader["Email"].ToString()
            };
        }
    }
}
