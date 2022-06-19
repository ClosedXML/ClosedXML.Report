using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace ClosedXML.Report.Utils
{
    public static class StringExtensions
    {
        public static string ReplaceFirst(this string source, string find, string replace)
        {
            int Place = source.IndexOf(find, StringComparison.Ordinal);
            string result = source.Remove(Place, find.Length).Insert(Place, replace);
            return result;
        }

        public static string ReplaceLast(this string source, string find, string replace)
        {
            int place = source.LastIndexOf(find, StringComparison.Ordinal);

            if (place == -1)
                return source;

            string result = source.Remove(place, find.Length).Insert(place, replace);
            return result;
        }

        /// <summary>
        /// Matching all capital letters in the input and separate them with spaces to form a sentence.
        /// If the input is an abbreviation text, no space will be added and returns the same input.
        /// </summary>
        /// <example>
        /// input : HelloWorld
        /// output : Hello World
        /// </example>
        /// <example>
        /// input : BBC
        /// output : BBC
        /// </example>
        /// <param name="input" />
        /// <returns/>
        public static string ToSentence(this string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;
            //return as is if the input is just an abbreviation
            if (Regex.Match(input, "[0-9A-Z]+$").Success)
                return input;
            //add a space before each capital letter, but not the first one.
            var result = Regex.Replace(input, "(\\B[A-Z])", " $1");
            return result;
        }

        /// <summary>
        /// Return last "howMany" characters from "input" string.
        /// </summary>
        /// <param name="input">Input string</param>
        /// <param name="howMany">Characters count to return</param>
        /// <returns></returns>
        public static string GetLast(this string input, int howMany)
        {
            if (string.IsNullOrEmpty(input)) return string.Empty;
            var value = input.Trim();
            return howMany >= value.Length ? value : value.Substring(value.Length - howMany);
        }

        /// <summary>
        /// Return first "howMany" characters from "input" string.
        /// </summary>
        /// <param name="input">Input string</param>
        /// <param name="howMany">Characters count to return. If howMany is negative then removed right chars.</param>
        /// <returns></returns>
        public static string GetFirst(this string input, int howMany)
        {
            if (string.IsNullOrEmpty(input)) return string.Empty;
            var value = input.Trim();
            if (howMany < 0)
                howMany = input.Length + howMany;
            return howMany >= value.Length ? value : input.Substring(0, howMany);
        }
         
        private static readonly Regex IsNumberRegex = new Regex(@"^[+-]?\d+$", RegexOptions.IgnoreCase);
        public static bool IsNumber(this string input)
        {
            var match = IsNumberRegex.Match(input);
            return match.Success;
        }

        private static readonly Regex EmailRegex = new Regex(@"\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*", RegexOptions.IgnoreCase);
        public static bool IsEmail(this string input)
        {
            var match = EmailRegex.Match(input);
            return match.Success;
        }

        private static readonly Regex PhoneRegex = new Regex(@"^\+?(\d[\d-. ]+)?(\([\d-. ]+\))?[\d-. ]+\d$", RegexOptions.IgnoreCase);
        public static bool IsPhone(this string input)
        {
            var match = PhoneRegex.Match(input);
            return match.Success;
        }

        private static readonly Regex IntRegex = new Regex(@"[+-]?\d+", RegexOptions.IgnoreCase);
        public static int ExtractNumber(this string input, int def)
        {
            if (string.IsNullOrEmpty(input)) return 0;

            var match = IntRegex.Match(input);
            return match.Success ? int.Parse(match.Value) : def;
        }

        /// <summary>
        /// Checks string object's value to array of string values
        /// </summary>
        /// <param name="value">Input string</param>
        /// <param name="stringValues">Array of string values to compare</param>
        /// <returns>Return true if any string value matches</returns>
        public static bool In(this string value, params string[] stringValues)
        {
            return stringValues.Any(otherValue => String.CompareOrdinal(value, otherValue) == 0);
        }

        public static bool IsNullOrWhiteSpace(this string input)
        {
            return string.IsNullOrEmpty(input) || input.Trim() == string.Empty;
        }

        /// <summary>
        /// Formats the string according to the specified mask
        /// </summary>
        /// <param name="input">The input string.</param>
        /// <param name="mask">The mask for formatting. Like "A##-##-T-###Z"</param>
        /// <returns>The formatted string</returns>
        public static string FormatWithMask(this string input, string mask)
        {
            if (input.IsNullOrWhiteSpace()) return input;
            var output = string.Empty;
            var index = 0;
            foreach (var m in mask)
            {
                if (m == '#')
                {
                    if (index < input.Length)
                    {
                        output += input[index];
                        index++;
                    }
                }
                else
                    output += m;
            }
            return output;
        }

        public static string Format(this string stringFormat, params object[] pars)
        {
            return string.Format(stringFormat, pars);
        }
    }
}
