using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XmlConverterJaarboek
{
    class Queries
    {
        public static string DOCTORS_FOR_INEXTENSO = 
            "SELECT * FROM [annuaire 2017] WHERE InExtenso = @inextenso AND CSouMS = @csms ORDER BY NOM, Institution";
        public static string DOCTORS_FOR_INEXTENSO_PERPROVINCE =
            "SELECT * FROM [annuaire 2017] WHERE InExtenso = @inextenso AND ProvArrond IS NOT NULL " +
            "ORDER BY ProvArrond, Poste, NOM";
    }
}
