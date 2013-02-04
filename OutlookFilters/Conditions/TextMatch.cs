using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Outlook;

namespace OutlookFilters.Conditions
{
    public class TextMatch : Condition
    {
        #region Enum
        [Flags]
        public enum TextFieldType
        {
            Subject = 1,
            Body = 2
        }
        public enum TextMatchType
        {
            Exact,
            Contains,
            Begin,
            End,
            RegEx
        }
        #endregion

        #region Properties
        public string        SearchText  { get; set; }
        public TextMatchType MatchMethod { get; set; }
        public TextFieldType FieldType { get; set; }
        #endregion

        #region Constructor
        public TextMatch()
        {
            SearchText = String.Empty;
            MatchMethod = TextMatchType.Exact;
            FieldType = TextFieldType.Body;
        }
        #endregion

        #region Condition Implementation
        public override bool Evaluate(MailItem item)
        {
            if (item == null)
                return false;

            try
            {
                var fields = new List<string>();
                if (FieldType.HasFlag(TextFieldType.Subject))
                {
                    fields.Add(item.Subject);
                }
                if (FieldType.HasFlag(TextFieldType.Body))
                {
                    fields.Add(item.Body);
                }
                
                switch(MatchMethod)
                {
                    case TextMatchType.Contains:
                        return fields.Any(t => t.Contains(SearchText));
                    case TextMatchType.Exact:
                        return fields.Any(t => t.Equals(SearchText));
                    case TextMatchType.RegEx:
                        return fields.Any(t => (new Regex(SearchText)).IsMatch(t));
                    case TextMatchType.Begin:
                        return fields.Any(t => t.StartsWith(SearchText));
                    case TextMatchType.End:
                        return fields.Any(t => t.EndsWith(SearchText));
                    default:
                        return false;
                }
            }
            catch (System.Exception ex)
            {
                return false;
            }
        }
        #endregion
    }
}
