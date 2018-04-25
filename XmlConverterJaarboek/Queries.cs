using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XmlConverterJaarboek
{
    class Queries
    {
        public static string DOCTORS_FOR_AFFILIATION = 
            "SELECT * FROM [annuaire 2017] WHERE Affiliation = @affiliation AND CSouMS = @csms ORDER BY NOM, Institution";
        public static string DOCTORS_FOR_AFFILIATION_PERPROVINCE =
            "SELECT * FROM [annuaire 2017] WHERE Affiliation = @affiliation AND ProvArrond IS NOT NULL " +
            "ORDER BY ProvArrond, Poste, NOM";
    }
}
