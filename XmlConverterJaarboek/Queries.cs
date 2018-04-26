using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XmlConverterJaarboek
{
    class Queries
    {
        public static string DOCTORS_FOR_INEXTENSO =
            "SELECT x.* FROM(SELECT *, IIF(SPECIALITE = \"ONCO_MED\", \"ONCO_MED\", Compétence2) AS Competence2Real, " +
            Properties.Settings.Default.InExtensoNew + " FROM [annuaire 2017]) AS x " +
            "WHERE x.InExtensoNew = @inextenso AND x.CSouMS = @csms ORDER BY x.NOM, x.Institution";
        public static string DOCTORS_FOR_INEXTENSO_PERPROVINCE =
            "SELECT x.* FROM(SELECT *, " + Properties.Settings.Default.InExtensoNew + " FROM [annuaire 2017]) AS x " +
            "WHERE x.InExtensoNew = @inextenso AND x.ProvArrond IS NOT NULL " +
            "ORDER BY x.ProvArrond, x.Poste, x.NOM";
        public static string DOCTORS_INTERNAL_FOR_COMPETENCE_PERPOSTAL =
            "SELECT x.NOM, x.PRENOM, x.Poste, x.Commune FROM(SELECT *, IIF(SPECIALITE = \"ONCO_MED\", \"ONCO_MED\", Compétence2) AS Competence2Real FROM [annuaire 2017]) AS x " +
            "WHERE (x.Compétence1 = @competence OR x.Competence2Real = @competence) AND x.Affiliation = \"MI\" AND x.Poste IS NOT NULL " +
            "GROUP BY x.NOM, x.PRENOM, x.Poste, x.Commune " +
            "ORDER BY x.Poste, x.NOM";
        public static string DOCTORS_ALPHABETIC =
            "SELECT x.NOM, x.PRENOM, x.NoINAMI, x.InExtensoNew, x.Compétence1, x.Competence2Real " +
            "FROM (SELECT *, IIF(SPECIALITE = \"ONCO_MED\", \"ONCO_MED\", Compétence2) AS Competence2Real, " + 
            Properties.Settings.Default.InExtensoNew + " FROM [annuaire 2017]) AS x " +
            "ORDER BY x.NOM, x.PRENOM";
    }
}
