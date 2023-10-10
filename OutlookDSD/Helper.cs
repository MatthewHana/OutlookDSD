using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace OutlookDSD
{
    public static class Helper
    {
        const string HEADER_TRANSPORT_SCHEMA = "http://schemas.microsoft.com/mapi/proptag/0x007D001E";
        const string HEADER_REGEX_PATTERN = @"^(?<header_key>[-A-Za-z0-9]+)(?<seperator>:[ \t]*)(?<header_value>([^\r\n]|\r\n[ \t]+)*)(?<terminator>\r\n)";

        public static ILookup<string, string> Email_GetHeaders(MailItem emailItem)
        {
            var headerString = (string)emailItem.PropertyAccessor.GetProperty(HEADER_TRANSPORT_SCHEMA);
            var headerMatches = Regex.Matches(headerString, HEADER_REGEX_PATTERN, RegexOptions.Multiline).Cast<Match>();
            return headerMatches.ToLookup(
                h => h.Groups["header_key"].Value,
                h => h.Groups["header_value"].Value
                );
        }

        public static string ListToSentence(List<string> list)
        {
            if (list.Count == 0)
            {
                return String.Empty;
            }
            else if (list.Count > 1)
            {
                return String.Join(", ", list.ToArray(), 0, list.Count - 1) + ", and " + list.LastOrDefault();
            }
            else
            {
                return list.First();
            }

        }
    }
}
