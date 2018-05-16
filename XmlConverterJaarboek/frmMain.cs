using System;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Xml;
using System.Collections.Generic;
using XmlConverterJaarboek.Entities;
using System.IO;
using System.ComponentModel;

namespace XmlConverterJaarboek
{
    public partial class frmMain : Form
    {
        Dictionary<string, string> provArrondMapping;
        OleDbConnection conn;
        string specialitiesFilePath = "";
        string alphabeticFilePath = "";

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

            var filePath = "db.mdb";
            if (!File.Exists(filePath))
            {
                DialogResult dr = ofdOpenFile.ShowDialog();

                if (dr == DialogResult.OK)
                {
                    filePath = ofdOpenFile.FileName;
                } else
                {
                    MessageBox.Show("Geen bestand gekozen, programma wordt afgesloten.");
                    this.Close();
                }

            }

            sfdDialog.Title = "Selecteer waar de lijst met specialiteiten moet worden opgeslagen";
            sfdDialog.FileName = "Lijst_specialiteiten.xml";
            sfdDialog.Filter = "eXtensible Markup Language (*.xml)|*.xml";
            if (sfdDialog.ShowDialog() == DialogResult.OK)
            {
                specialitiesFilePath = sfdDialog.FileName;
            }
            else
            {
                MessageBox.Show("Geen bestand gekozen, programma wordt afgesloten.");
                this.Close();
            }

            sfdDialog.Title = "Selecteer waar de alfabetische lijst moet worden opgeslagen";
            sfdDialog.FileName = "Lijst_alfabetisch.xml";
            sfdDialog.Filter = "eXtensible Markup Language (*.xml)|*.xml";
            if (sfdDialog.ShowDialog() == DialogResult.OK)
            {
                alphabeticFilePath = sfdDialog.FileName;
            }
            else
            {
                MessageBox.Show("Geen bestand gekozen, programma wordt afgesloten.");
                this.Close();
            }

            InputDialog id = new InputDialog();
            id.lblText.Text = "Vul de naam in van de tabel met alle artsen. Voorbeeld: annuaire 2017 (deze waarde kan elk jaar veranderen, voor 2018 kan dit \"annuaire 2018\" zijn)";
            
            if (id.ShowDialog() == DialogResult.OK)
            {
                Queries.TABLE_NAME = id.txtValue.Text;
            }
            else
            {
                Application.Exit();
            }
            
            conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" + filePath);
            //using (OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db.mdb"))
            //using (OleDbConnection conn = new OleDbConnection(@"Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=db.mdb;"))
            conn.Open();
            backgroundWorker.RunWorkerAsync();
            backgroundWorker.RunWorkerCompleted += BackgroundWorker_RunWorkerCompleted;
            backgroundWorker.ProgressChanged += BackgroundWorker_ProgressChanged;
            //Close();
        }

        private void BackgroundWorker_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            progressBar1.Value = Math.Min(100, Math.Max(0, e.ProgressPercentage));
            lblCurrent.Text = (string) e.UserState;
        }

        private void BackgroundWorker_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            conn.Close();
            this.Close();
        }

        private void ConvertDoctorsInternalGrouped(OleDbConnection conn, XmlWriter writer)
        {
            writer.WriteElementString("_10_Titel", "Membres de l’union professionelle titulaire d’un titre professionel particulier/Leden van de beroepsvereniging houders van bijzondere beroepstitel" + Characters.PARAGRAPH_SEP);

            for (int i = 0; i < Properties.Settings.Default.InternalOrder.Count; i++)
            {
                var currentInternalOrder = Properties.Settings.Default.InternalOrder[i];
                var currentInternalNames = Properties.Settings.Default.InternalNames[i].Split(',');

                List<SimpleDoctor> doctors = GetDoctorsInternalForCompetence(conn, currentInternalOrder);

                if (doctors.Count > 0)
                {
                    writer.WriteElementString("_11_TitelF", currentInternalNames[0] + Characters.PARAGRAPH_SEP);
                    writer.WriteElementString("_11_TitelN", currentInternalNames[1] + Characters.PARAGRAPH_SEP);
                }

                string previousPostalCode = null;

                foreach (SimpleDoctor doctor in doctors)
                {
                    if (previousPostalCode == null || previousPostalCode != doctor.PostalCode)
                    {
                        writer.WriteElementString("_12_Gegevens", doctor.PostalCode + " " + doctor.Town + "\t" + doctor.LastName.ToUpper() + " " + doctor.FirstName.ToCapitalized() + Characters.PARAGRAPH_SEP);
                    }
                    else
                    {
                        writer.WriteElementString("_12_Gegevens", "\t" + doctor.LastName.ToUpper() + " " + doctor.FirstName.ToCapitalized() + Characters.PARAGRAPH_SEP);
                    }

                    previousPostalCode = doctor.PostalCode;
                }

            }

            writer.WriteElementString("_09_Einde", Characters.PARAGRAPH_SEP);
        }

        private void ConvertDoctorsForInExtensoPerProvince(OleDbConnection conn, string inextenso, XmlWriter writer)
        {
            writer.WriteElementString("_05_Titel", "Liste par province/Lijst per provincie" + Characters.PARAGRAPH_SEP);

            Dictionary<string, List<SimpleDoctor>> docsPerProvince = GetDoctorsForInExtensoPerProvince(conn, inextenso);

            foreach (string orderedProvince in Properties.Settings.Default.ProvArrondMapOrder)
            {
                List<SimpleDoctor> doctors = docsPerProvince[orderedProvince];

                if (doctors.Count > 0)
                    writer.WriteElementString("_06_Provincie", orderedProvince.ToUpper() + Characters.PARAGRAPH_SEP);

                string previousPostalCode = null;
                SimpleDoctor lastDoctor = null;
                foreach (SimpleDoctor doctor in doctors)
                {
                    if (previousPostalCode == null || previousPostalCode != doctor.PostalCode)
                    {
                        writer.WriteElementString("_07_Stad", doctor.PostalCode + " " + doctor.Town.ToUpper() + Characters.PARAGRAPH_SEP);
                        previousPostalCode = doctor.PostalCode;
                        lastDoctor = null;
                    }

                    bool shouldWriteDoctor = true;
                    if (lastDoctor != null && lastDoctor.IsSameDoctor(doctor.FirstName, doctor.LastName, doctor.INAMI))
                        shouldWriteDoctor = false;

                    if (shouldWriteDoctor)
                    {
                        writer.WriteElementString("_08_Gegevens", doctor.LastName.ToUpper() + " " + doctor.FirstName
                        + (doctor.CSouMS == "CS" ? Characters.FIXED_SPACE + "(*)" : "") + Characters.PARAGRAPH_SEP);
                    }

                    lastDoctor = doctor;
                }

            }
        }

        private void ConvertDoctorsForInExtenso(OleDbConnection conn, string inextenso, string csms, XmlWriter writer)
        {
            bool titleDone = false;
            bool firstDoctor = true;

            foreach (Doctor doctor in GetDoctorsForInExtenso(conn, inextenso, csms))
            {
                if (csms == "CS" && !titleDone)
                {
                    writer.WriteElementString("_04_Titel", "CANDIDATS SPECIALISTES•KANDIDATEN-SPECIALISTEN" + Characters.PARAGRAPH_SEP);
                    titleDone = true;
                }

                if (csms != "CS" && firstDoctor)
                {
                    writer.WriteElementString("_02_Naam1", doctor.GetFormattedDetails() + Characters.PARAGRAPH_SEP);
                    firstDoctor = false;
                } else
                {
                    writer.WriteElementString("_02_Naam", doctor.GetFormattedDetails() + Characters.PARAGRAPH_SEP);
                }

                bool first = true;
                foreach (ContactDetails contactDetails in doctor.ContactDetails)
                {
                    var elementName = first ? "_03_Gegevens1" : "_03_Gegevens2";
                    var details = contactDetails.GetFormattedDetails();

                    writer.WriteElementString(elementName, details + (details.Trim().Length == 0 ? "" : Characters.PARAGRAPH_SEP));

                    first = false;
                }
            }
        }

        private void ConvertDoctorsCandidates(OleDbConnection conn, XmlWriter writer, BackgroundWorker worker)
        {
            bool firstDoctor = true;

            worker.ReportProgress(0, "Candidates");
            var i = 0;
            var candidates = GetDoctorsCandidates(conn);
            foreach (Doctor doctor in candidates)
            {
                var currentProgress = (int)Math.Round(((float)i / (float)(candidates.Count - 1) * 100f));
                backgroundWorker.ReportProgress(currentProgress, "Candidate " + i.ToString());

                if (firstDoctor)
                {
                    writer.WriteElementString("_02_Naam1", doctor.GetFormattedDetails() + Characters.PARAGRAPH_SEP);
                    firstDoctor = false;
                }
                else
                {
                    writer.WriteElementString("_02_Naam", doctor.GetFormattedDetails() + Characters.PARAGRAPH_SEP);
                }

                bool first = true;
                foreach (ContactDetails contactDetails in doctor.ContactDetails)
                {
                    var elementName = first ? "_03_Gegevens1" : "_03_Gegevens2";
                    var details = contactDetails.GetFormattedDetails();

                    writer.WriteElementString(elementName, details + (details.Trim().Length == 0 ? "" : Characters.PARAGRAPH_SEP));

                    first = false;
                }

                i++;
            }
        }

        private List<Doctor> GetDoctorsForInExtenso(OleDbConnection conn, string inextenso, string csms)
        {
            var command = conn.CreateCommand();
            command.CommandText = Queries.DOCTORS_FOR_INEXTENSO();
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
                        InExtenso = reader["InExtensoNew"].ToString(),
                        Competence1 = reader["Compétence1"].ToString(),
                        Competence2 = reader["Competence2Real"].ToString()
                    };

                    newDoctor.ContactDetails.Add(CreateContactDetails(reader));

                    doctorList.Add(newDoctor);
                }
            }
            reader.Close();

            return doctorList;
        }

        private List<Doctor> GetDoctorsCandidates(OleDbConnection conn)
        {
            var command = conn.CreateCommand();
            command.CommandText = Queries.DOCTORS_CANDIDATES();
            command.Parameters.AddRange(new OleDbParameter[]
            {
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
                        InExtenso = reader["InExtensoNew"].ToString(),
                        Competence1 = reader["Compétence1"].ToString(),
                        Competence2 = reader["Competence2Real"].ToString()
                    };

                    newDoctor.ContactDetails.Add(CreateContactDetails(reader));

                    doctorList.Add(newDoctor);
                }
            }
            reader.Close();

            return doctorList;
        }

        private List<SimpleDoctor> GetDoctorsInternalForCompetence(OleDbConnection conn, string competence)
        {
            var command = conn.CreateCommand();
            command.CommandText = Queries.DOCTORS_INTERNAL_FOR_COMPETENCE_PERPOSTAL();
            command.Parameters.AddRange(new OleDbParameter[]
            {
               new OleDbParameter("@competence", competence)
            });

            List<SimpleDoctor> doctorList = new List<SimpleDoctor>();
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

                doctorList.Add(newDoctor);
            }
            reader.Close();

            return doctorList;
        }

        private List<AlphabeticDoctor> GetDoctorsAlphabetic(OleDbConnection conn)
        {
            var command = conn.CreateCommand();
            command.CommandText = Queries.DOCTORS_ALPHABETIC();

            List<AlphabeticDoctor> doctors = new List<AlphabeticDoctor>();
            var reader = command.ExecuteReader();
            AlphabeticDoctor newDoctor = null;
            while (reader.Read())
            {
                if (newDoctor != null
                    && newDoctor.IsSameDoctor(
                        reader["PRENOM"].ToString(),
                        reader["NOM"].ToString(),
                        reader["NoINAMI"].ToString()))
                {
                    var containsExtenso = false;
                    foreach (ExtensoDetails ed in newDoctor.Extensos)
                    {
                        if (ed.Name.Equals(reader["InExtensoNew"].ToString()))
                        {
                            containsExtenso = true;
                        }
                    }

                    if (!containsExtenso)
                    {
                        List<string> competences = new List<string>();
                        competences.Add(reader["Compétence1"].ToString());
                        if (reader["Competence2Real"].ToString() != "") competences.Add(reader["Competence2Real"].ToString());

                        string competenceString = string.Join("-", competences.ToArray());

                        var newExtenso = new ExtensoDetails
                        {
                            Name = reader["InExtensoNew"].ToString(),
                            Competences = competenceString
                        };

                        newDoctor.Extensos.Add(newExtenso);
                    }
                }
                else
                {
                    newDoctor = new AlphabeticDoctor
                    {
                        FirstName = reader["PRENOM"].ToString(),
                        LastName = reader["NOM"].ToString(),
                        INAMI = reader["NoINAMI"].ToString()
                    };

                    List<string> competences = new List<string>();
                    competences.Add(reader["Compétence1"].ToString());
                    if (reader["Competence2Real"].ToString() != "") competences.Add(reader["Competence2Real"].ToString());

                    string competenceString = string.Join("-", competences.ToArray());

                    var newExtenso = new ExtensoDetails
                    {
                        Name = reader["InExtensoNew"].ToString(),
                        Competences = competenceString
                    };

                    newDoctor.Extensos.Add(newExtenso);

                    doctors.Add(newDoctor);
                }
            }
            reader.Close();

            return doctors;
        }

        private Dictionary<string, List<SimpleDoctor>> GetDoctorsForInExtensoPerProvince(OleDbConnection conn, string inextenso)
        {
            var command = conn.CreateCommand();
            command.CommandText = Queries.DOCTORS_FOR_INEXTENSO_PERPROVINCE();
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
                    Town = reader["Commune"].ToString(),
                    CSouMS = reader["CSouMS"].ToString(),
                    INAMI = reader["NoINAMI"].ToString()
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

        private void backgroundWorker_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            var writer = XmlWriter.Create(specialitiesFilePath);
            writer.WriteStartDocument(true);
            writer.WriteStartElement("Root");

            //Properties.Settings.Default.SpecialisationOrder.Count;

            var i = 0;
            foreach (string specialisation in Properties.Settings.Default.SpecialisationOrder)
            {
                var currentProgress = (int)Math.Round(((float)i / (float)(Properties.Settings.Default.SpecialisationOrder.Count - 1) * 100f));
                backgroundWorker.ReportProgress(currentProgress, specialisation + " ...");

                if (GetDoctorsForInExtenso(conn, specialisation, "MS").Count == 0)
                {
                    MessageBox.Show("No doctors for specialisation: " + specialisation);
                }

                writer.WriteElementString("_09_Einde", Characters.PARAGRAPH_SEP);
                backgroundWorker.ReportProgress(currentProgress, specialisation + " 1/3 MS");
                ConvertDoctorsForInExtenso(conn, specialisation, "MS", writer);
                backgroundWorker.ReportProgress(currentProgress, specialisation + " 2/3 CS");
                ConvertDoctorsForInExtenso(conn, specialisation, "CS", writer);
                writer.WriteElementString("_09_Einde", Characters.PARAGRAPH_SEP);

                if (specialisation.Equals("MED. INTERNE"))
                {
                    backgroundWorker.ReportProgress(currentProgress, specialisation + " 2/3B Internal");
                    ConvertDoctorsInternalGrouped(conn, writer);
                }

                backgroundWorker.ReportProgress(currentProgress, specialisation + " 3/3 Provinces");
                ConvertDoctorsForInExtensoPerProvince(conn, specialisation, writer);

                i++;
            }

            //Candidates specialistes
            ConvertDoctorsCandidates(conn, writer, backgroundWorker);            

            writer.WriteString(Characters.PARAGRAPH_SEP);
            writer.WriteEndElement();
            writer.Flush();

            writer = XmlWriter.Create(alphabeticFilePath);
            writer.WriteStartElement("Root");

            backgroundWorker.ReportProgress(100, "Alphabetic list");

            List<AlphabeticDoctor> doctors = GetDoctorsAlphabetic(conn);
            i = 0;
            string currentLetter = null;
            foreach (AlphabeticDoctor doctor in doctors)
            {
                var firstLetter = doctor.LastName.Substring(0, 1).ToUpper();

                if (currentLetter == null || currentLetter != firstLetter)
                {
                    if (currentLetter != null)
                    {
                        writer.WriteEndElement();
                    }

                    currentLetter = firstLetter;

                    writer.WriteElementString("nr", currentLetter + Characters.PARAGRAPH_SEP);
                    writer.WriteStartElement("lijst");
                }

                bool first = true;
                foreach (ExtensoDetails extenso in doctor.Extensos)
                {
                    var extensoAndNum = extenso.Name + " (" + 
                        (Properties.Settings.Default.SpecialisationOrder.IndexOf(extenso.Name) + 1).ToString() + ")\t" 
                        + extenso.Competences + Characters.PARAGRAPH_SEP;

                    if (first)
                    {
                        writer.WriteString(doctor.LastName + " " + doctor.FirstName + "\t" + extensoAndNum);
                    }
                    else
                    {
                        writer.WriteString("\t" + extensoAndNum);
                    }

                    first = false;
                }

                i++;
            }

            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.Flush();
        }
    }
}
