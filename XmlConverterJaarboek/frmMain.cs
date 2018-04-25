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
        Dictionary<string, string> provArrondMapping;

        public frmMain()
        {
            InitializeComponent();
        }

        private void FormLoad(object sender, EventArgs e)
        {
            provArrondMapping = new Dictionary<string, string>();
            
            foreach (string mapping in Properties.Settings.Default.ProvArrondMapping)
            {
                var province = mapping.Split('=')[1];
                var values = mapping.Split('=')[0].Split(',');
                
                foreach (string value in values)
                {
                    provArrondMapping.Add(value, province);
                }
            }

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

            foreach (string specialisation in Properties.Settings.Default.SpecialisationOrder)
            {
                if (GetDoctorsForInExtenso(conn, specialisation, "MS").Count == 0)
                {
                    MessageBox.Show("No doctors for specialisation: " + specialisation);
                }

                writer.WriteElementString("_09_Einde", "");
                ConvertDoctorsForInExtenso(conn, specialisation, "MS", writer);
                ConvertDoctorsForInExtenso(conn, specialisation, "CS", writer);
                writer.WriteElementString("_09_Einde", "");
                ConvertDoctorsForInExtensoPerProvince(conn, specialisation, writer);
            }

            writer.WriteEndElement();
            writer.Flush();
        }

        private void ConvertDoctorsForInExtensoPerProvince(OleDbConnection conn, string inextenso, XmlWriter writer)
        {
            writer.WriteElementString("_05_Titel", "Liste par province/Lijst per provincie");

            Dictionary<string, List<SimpleDoctor>> docsPerProvince = GetDoctorsForInExtensoPerProvince(conn, inextenso);

            foreach (string orderedProvince in Properties.Settings.Default.ProvArrondMapOrder)
            {
                writer.WriteElementString("_06_Provincie", orderedProvince.ToUpper());

                List<SimpleDoctor> doctors = docsPerProvince[orderedProvince];
                string previousPostalCode = null;

                foreach (SimpleDoctor doctor in doctors)
                {
                    if (previousPostalCode == null || previousPostalCode != doctor.PostalCode)
                    {
                        writer.WriteElementString("_07_Stad", doctor.PostalCode + " " + doctor.Town.ToUpper());
                        previousPostalCode = doctor.PostalCode;
                    }

                    writer.WriteElementString("_08_Gegevens", doctor.LastName.ToUpper() + " " + doctor.FirstName);
                }

            }


        }

        private void ConvertDoctorsForInExtenso(OleDbConnection conn, string inextenso, string csms, XmlWriter writer)
        {
            bool titleDone = false;
            foreach (Doctor doctor in GetDoctorsForInExtenso(conn, inextenso, csms))
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

        private List<Doctor> GetDoctorsForInExtenso(OleDbConnection conn, string inextenso, string csms)
        {
            var command = conn.CreateCommand();
            command.CommandText = Queries.DOCTORS_FOR_INEXTENSO;
            command.Parameters.AddRange(new OleDbParameter[]
            {
               new OleDbParameter("@inextenso", inextenso),
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
                        InExtenso = reader["InExtenso"].ToString(),
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

        private Dictionary<string, List<SimpleDoctor>> GetDoctorsForInExtensoPerProvince(OleDbConnection conn, string inextenso)
        {
            var command = conn.CreateCommand();
            command.CommandText = Queries.DOCTORS_FOR_INEXTENSO_PERPROVINCE;
            command.Parameters.AddRange(new OleDbParameter[]
            {
               new OleDbParameter("@inextenso", inextenso)
            });

            Dictionary<string, List<SimpleDoctor>> doctorsPerProvince = new Dictionary<string, List<SimpleDoctor>>(); ;

            foreach (string orderedProvince in Properties.Settings.Default.ProvArrondMapOrder)
            {
                doctorsPerProvince.Add(orderedProvince, new List<SimpleDoctor>());
            }

            var reader = command.ExecuteReader();
            while (reader.Read())
            {
                SimpleDoctor newDoctor = new SimpleDoctor
                {
                    FirstName = reader["PRENOM"].ToString(),
                    LastName = reader["NOM"].ToString(),
                    PostalCode = reader["Poste"].ToString(),
                    Town = reader["Commune"].ToString()
                };

                if (!provArrondMapping.ContainsKey(reader["ProvArrond"].ToString()))
                {
                    MessageBox.Show("Unmapped province for ProvArrond " + reader["ProvArrond"].ToString());
                }
                else
                {
                    var provinceName = provArrondMapping[reader["ProvArrond"].ToString()];

                    if (!doctorsPerProvince.ContainsKey(provinceName))
                    {
                        MessageBox.Show(provinceName);
                    }

                    doctorsPerProvince[provinceName].Add(newDoctor);

                }
            }
            reader.Close();

            return doctorsPerProvince;
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
