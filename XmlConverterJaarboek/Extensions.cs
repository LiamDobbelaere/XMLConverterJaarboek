using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XmlConverterJaarboek
{
    static class Extensions
    {
        public static string ToCapitalized(this string str)
        {
            return str.Substring(0, 1).ToUpper() + str.Substring(1);
        }
    }
}
