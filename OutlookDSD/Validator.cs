using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace OutlookDSD
{
    public class Validator
    {
        public const string RESULT_NONE = "none";
        public const string RESULT_PASS = "pass";
        public const string RESULT_FAIL = "fail";
        public const string RESULT_ERROR = "error";

        public const string KEY_DKIM = "dkim";
        public const string KEY_SPF = "spf";
        public const string KEY_DMARC = "dmarc";

        public const string HEADER_AUTHRESULTS = "Authentication-Results";
        public const string HEADER_RECEIVED = "Received";

        private readonly MailItem emailitem;
        private readonly List<Dictionary<string, string>> authResults;
        private readonly Dictionary<string, string[]> results;
#pragma warning disable IDE0044 // Add readonly modifier
        private ILookup<string, string> emailHeaders;
#pragma warning restore IDE0044 // Add readonly modifier

        public bool isSent = true;

        public Validator(MailItem mailItem)
        {
            if (mailItem != null)
            {
                emailitem = mailItem;
                emailHeaders = Helper.Email_GetHeaders(emailitem);
                isSent = Email_IsSent();
                authResults = Email_ParseAuthenticationResults();
            }
            results = Email_Parse();
        }

        public Dictionary<string, string[]> Results()
        {
            return results;
        }

        private Dictionary<string, string[]> Email_Parse()
        {
            Dictionary<string, string[]> result = new Dictionary<string, string[]>();
            if (authResults == null)
            {
                result[KEY_DKIM] = new string[] { RESULT_ERROR, String.Empty };
                result[KEY_SPF] = new string[] { RESULT_ERROR, String.Empty };
                result[KEY_DMARC] = new string[] { RESULT_ERROR, String.Empty };
                return result;
            }

            result[KEY_DKIM] = Parse_DKIM();
            result[KEY_SPF] = Parse_SPF();
            result[KEY_DMARC] = Parse_DMARC();
            return result;
        }

        private string[] Parse_DKIM()
        {
            return Parse_Mechanism(KEY_DKIM);
        }

        private string[] Parse_SPF()
        {
            return Parse_Mechanism(KEY_SPF);
        }

        private string[] Parse_DMARC()
        {
            return Parse_Mechanism(KEY_DMARC);
        }

        private string[] Parse_Mechanism(string keyName)
        {
            string resultValue = RESULT_NONE;
            string details = "";
            // Iterare through each item in the AuthResults array
            foreach (Dictionary<string, string> authResultSegment in authResults)
            {
                if (!authResultSegment.ContainsKey(keyName))
                {
                    continue;
                }

                // Try to get the result from this dictionary 
                authResultSegment.TryGetValue(keyName, out string mechnaismResultValue);
                mechnaismResultValue = mechnaismResultValue.ToLower();

                if (mechnaismResultValue == "pass")
                {
                    resultValue = RESULT_PASS;
                }
                else if (mechnaismResultValue == "fail")
                {
                    resultValue = RESULT_FAIL;
                }
                else
                {
                    resultValue = RESULT_ERROR;
                }

                authResultSegment.TryGetValue("FULL", out details);

                break;
            }
            return new string[] { resultValue, details };
        }

        private bool Email_IsSent()
        {
            // If there are no Received headers then it's a sent email
            return !emailHeaders.Contains(HEADER_RECEIVED);

        }
        private List<Dictionary<string, string>> Email_ParseAuthenticationResults()
        {
            List<Dictionary<string, string>> cleanAuthResultsSegments = new List<Dictionary<string, string>>();

            // Return the empty list if there are no Authentication Results headers
            if (!emailHeaders.Contains(HEADER_AUTHRESULTS))
            {
                return cleanAuthResultsSegments;
            }

            string[] authResultsArray = emailHeaders[HEADER_AUTHRESULTS].ToArray();

            string cleanPattrn = @"\(([^)]+)\)";

            foreach (string authResult in authResultsArray)
            {
                // Remove any parenthesis and everything in between 
                string authResultClean = Regex.Replace(authResult, cleanPattrn, "");

                // Split each segment of the Authentication Results header by semicolon
                string[] authResultSegments = authResultClean.Split(";"[0]);
#pragma warning disable IDE0028 // Simplify collection initialization
                Dictionary<string, string> valuePairs = new Dictionary<string, string>();
#pragma warning restore IDE0028 // Simplify collection initialization

                // Add the full details to the valuePairs under the key "FULL"
                valuePairs.Add("FULL", authResult);

                for (int i = 0; i < authResultSegments.Length; i++)
                {
                    // Go through and remove whitespaces
                    string authResultSegment = authResultSegments[i].Trim();

                    // Split on each white space char
                    string[] lineSegments = authResultSegment.Split(" "[0]);


                    // Split each segment on the = sign if it exists
                    for (int j = 0; j < lineSegments.Length; j++)
                    {
                        string lineSegment = lineSegments[j].Trim();
                        string key;
                        string value;

                        // If the segment has no equals sign then add the entire string to the valuepair
                        int equalPos = lineSegment.IndexOf('=');
                        if (equalPos == -1)
                        {
                            key = lineSegment;
                            value = "";
                        }
                        // Otherwise split the segment at the = sign
                        else
                        {
                            key = lineSegment.Substring(0, equalPos);
                            value = lineSegment.Substring(equalPos + 1);
                        }

                        // If the key is empty then skip it.
                        if (key.Trim().Length == 0)
                        {
                            continue;
                        }

                        // If we don't an entry with this key then set it and continue with the rest of the loop
                        if (!valuePairs.ContainsKey(key))
                        {
                            valuePairs.Add(key, value);
                            continue;
                        }

                        // if we do, then let's check if they match
                        valuePairs.TryGetValue(key, out string existingValue);

                        // If they do match, then just move on
                        if (existingValue == value)
                        {
                            continue;
                        }

                        // Fail results take priority
                        if(value == "fail")
                        {
                            valuePairs[key] = value;
                        }
                        
                        //Console.WriteLine("Duplicate AuthResult key for " + key + " with different values.");
                        //Console.WriteLine("Current Value: " + existingValue);
                        //Console.WriteLine("New Value: " + value);
                    }
                    cleanAuthResultsSegments.Add(valuePairs);
                }
            }

            return cleanAuthResultsSegments;
        }

        public MailItem GetMailItem()
        {
            return emailitem;
        }

    }
}
