using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace Com.Ena.Timesheet.XlUtil
{
    public static class Text
    {
        /// <summary>
        /// Replace non-ASCII characters with spaces.
        /// </summary>
        public static string ReplaceNonAscii(string s)
        {
            if (s == null) return null;
            // Replace all non-ASCII characters with a space
            return Regex.Replace(s, @"[^\x00-\x7F]", " ");
        }

        /// <summary>
        /// Replace all quotes in a string.
        /// </summary>
        public static string Unquote(string s)
        {
            if (s == null) return null;
            // Replace all double quotes with an empty string
            s = s.Replace("\"", "");
            // Replace all single quotes with a backtick
            s = s.Replace("'", "`");
            return s;
        }

        public static KeyValuePair<int, string> ParseInt(string name, string s)
        {
            if (int.TryParse(s, out int value))
            {
                return new KeyValuePair<int, string>(value, "");
            }
            else
            {
                return new KeyValuePair<int, string>(0, $"{name} is not a number: {s}. ");
            }
        }

        public static KeyValuePair<TimeSpan, string> ParseXlTime(string name, string s)
        {
            // "Sun Dec 31 09:30:00 EST 1899"
            var parts = s?.Split(' ');
            if (parts != null && parts.Length == 6)
            {
                return ParseTime(name, parts[3]);
            }
            else
            {
                return new KeyValuePair<TimeSpan, string>(TimeSpan.Zero, $"{name} is not a time: {s}. ");
            }
        }

        public static KeyValuePair<TimeSpan, string> ParseTime(string name, string s)
        {
            if (TimeSpan.TryParse(s, out var time))
            {
                return new KeyValuePair<TimeSpan, string>(time, "");
            }
            else
            {
                return new KeyValuePair<TimeSpan, string>(TimeSpan.Zero, $"{name} is not a time: {s}. ");
            }
        }

        public static KeyValuePair<float, string> ParseFloat(string name, string s)
        {
            if (float.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out float value))
            {
                return new KeyValuePair<float, string>(value, "");
            }
            else
            {
                return new KeyValuePair<float, string>(0.0f, $"{name} is not a number: {s}. ");
            }
        }

        public static string StackFormatter1(Exception e)
        {
            var sb = new StringBuilder();
            sb.AppendLine(e.Message);
            foreach (var ste in e.StackTrace?.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries) ?? Array.Empty<string>())
            {
                sb.AppendLine(ste);
            }
            return sb.ToString();
        }

        public static string StackFormatter2(Exception e)
        {
            return e.ToString();
        }

        public static bool IsBlank(string s)
        {
            return string.IsNullOrWhiteSpace(s);
        }
    }
}