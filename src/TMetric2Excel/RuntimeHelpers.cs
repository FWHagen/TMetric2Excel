using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TMetric2Excel
{
    public static class RuntimeHelpers
    {
        public static bool IsEmpty(this string value)
        {
            return String.IsNullOrWhiteSpace(value);
        }
        public static bool IsNotEmpty(this string value)
        {
            return !String.IsNullOrWhiteSpace(value);
        }

        public static bool IsAlphaNumeric(this string str)
        {
            var regexItem = new System.Text.RegularExpressions.Regex("^[a-zA-Z0-9]*$");

            return (regexItem.IsMatch(str));
        }

        public static string ToAlpha(this string value)
        {
            if (value.IsEmpty())
                return value;
            return System.Text.RegularExpressions.Regex.Replace(value, "[^a-zA-Z]", String.Empty);
        }
        public static string ToAlphaNumeric(this string value)
        {
            if (value.IsEmpty())
                return value;
            return System.Text.RegularExpressions.Regex.Replace(value, "[^a-zA-Z0-9.]", String.Empty);
        }

        public static double ToRealOrZero(this string value, int decimalPlaces = 2)
        {
            double result = 0;
            if (Double.TryParse(value, out result))
                return Math.Round(result, decimalPlaces);

            return 0.00;
        }

        public static string ToIsoDate(this DateTime dateval, int places = 8)
        {
            string result = "";
            result = String.Format("{0:yyyyMMddHHmmss}", dateval);
            if (places < 9)
            {
                result = result.Substring(0, 8);
                if (places == 8)
                    return result;
                return result.Substring(8 - places);
            }
            if (places < 14)
                return result.Substring(0, places);

            return result;
        }

        public static bool ToOptimisticBool(this string value)
        {
            if (String.IsNullOrWhiteSpace(value))
                return true;
            value = value.ToLower();
            if (value == "f" || value == "false" || value == "n" || value == "no" || value == "0" || value == "off")
                return false;
            return true;
        }
        public static bool ToPessimisticBool(this string value)
        {
            if (String.IsNullOrWhiteSpace(value))
                return false;
            value = value.ToLower();
            if (value == "t" || value == "true" || value == "y" || value == "yes" || value == "1" || value == "on")
                return true;
            return false;
        }




        public static string ThisMethod([System.Runtime.CompilerServices.CallerMemberName] string remotemethod = "") { return remotemethod; }


        public static string SafeFilename(this string value)
        {
            if (value.IsEmpty())
                return value;
            return System.Text.RegularExpressions.Regex.Replace(value, "[^a-zA-Z0-9.-_]", String.Empty);
        }
    }
}
