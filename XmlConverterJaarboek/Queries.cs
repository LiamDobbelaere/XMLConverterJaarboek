using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XmlConverterJaarboek
{
    class Queries
    {
        public static string TABLE_NAME = "annuaire 2017";

        public static string DOCTORS_FOR_INEXTENSO() {
            return "SELECT x.* FROM(SELECT *, IIF(SPECIALITE = \"ONCO_MED\", \"ONCO_MED\", Compétence2) AS Competence2Real, " +
            Properties.Settings.Default.InExtensoNew + " FROM [" + TABLE_NAME + "]) AS x " +
            "WHERE x.InExtensoNew = @inextenso AND x.CSouMS = @csms ORDER BY x.NOM, x.PRENOM, x.Institution";
        }

        public static string DOCTORS_CANDIDATES()
        {
            return "SELECT x.* FROM(SELECT *, IIF(SPECIALITE = \"ONCO_MED\", \"ONCO_MED\", Compétence2) AS Competence2Real, " +
            Properties.Settings.Default.InExtensoNew + " FROM [" + TABLE_NAME + "]) AS x " +
            "WHERE x.CSouMS = \"CS\" ORDER BY x.NOM, x.PRENOM, x.Institution";
        }

        public static string DOCTORS_FOR_INEXTENSO_PERPROVINCE()
        {
            return "SELECT x.NOM, x.PRENOM, x.NoINAMI, x.Poste, x.Commune, x.CSouMS, x.ProvArrond, x.InExtensoNew FROM(SELECT *, " + Properties.Settings.Default.InExtensoNew + " FROM [" + TABLE_NAME + "]) AS x " +
            "WHERE x.InExtensoNew = @inextenso AND x.ProvArrond IS NOT NULL " +
            "GROUP BY x.PRENOM, x.NOM, x.NoINAMI, x.Poste, x.Commune, x.CSouMS, x.ProvArrond, x.InExtensoNew " +
            "ORDER BY x.ProvArrond, x.Poste, x.NOM, x.PRENOM";
        }

        public static string DOCTORS_INTERNAL_FOR_COMPETENCE_PERPOSTAL() {
            return "SELECT x.NOM, x.PRENOM, x.Poste, x.Commune FROM(SELECT *, IIF(SPECIALITE = \"ONCO_MED\", \"ONCO_MED\", Compétence2) AS Competence2Real FROM [" + TABLE_NAME + "]) AS x " +
            "WHERE (x.Compétence1 = @competence OR x.Competence2Real = @competence) AND x.Affiliation = \"MI\" AND x.Poste IS NOT NULL " +
            "GROUP BY x.NOM, x.PRENOM, x.Poste, x.Commune " +
            "ORDER BY x.Poste, x.NOM, x.PRENOM";
        }

        public static string DOCTORS_ALPHABETIC()
        {
            return "SELECT x.NOM, x.PRENOM, x.NoINAMI, x.InExtensoNew, x.Compétence1, x.Competence2Real " +
            "FROM (SELECT *, IIF(SPECIALITE = \"ONCO_MED\", \"ONCO_MED\", Compétence2) AS Competence2Real, " +
            Properties.Settings.Default.InExtensoNew + " FROM [" + TABLE_NAME + "]) AS x " +
            "ORDER BY x.NOM, x.PRENOM";
        }
    }
}
