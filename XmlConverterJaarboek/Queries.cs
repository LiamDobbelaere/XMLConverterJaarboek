using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XmlConverterJaarboek
{
    class Queries
    {
        public static string DOCTORS_FOR_INEXTENSO = 
            "SELECT x.* FROM(SELECT *, " + Properties.Settings.Default.InExtensoNew + " FROM [annuaire 2017]) AS x " +
            "WHERE x.InExtensoNew = @inextenso AND x.CSouMS = @csms ORDER BY x.NOM, x.Institution";
        public static string DOCTORS_FOR_INEXTENSO_PERPROVINCE =
            "SELECT x.* FROM(SELECT *, " + Properties.Settings.Default.InExtensoNew + " FROM [annuaire 2017]) AS x " +
            "WHERE x.InExtensoNew = @inextenso AND x.ProvArrond IS NOT NULL " +
            "ORDER BY x.ProvArrond, x.Poste, x.NOM";
    }
}
