using System;
using System.Collections.Generic;
using System.Text;

namespace Frends.Community.PdfFromTemplate
{
    static class Extensions
    {
        public static TEnum ConvertEnum<TEnum>(this Enum source)
        {
            return (TEnum)Enum.Parse(typeof(TEnum), source.ToString(), true);
        }
    }
}
