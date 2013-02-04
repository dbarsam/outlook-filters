using System.Collections.Generic;
using System.Linq;
using System.Xml.Serialization;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookFilters.Filters
{
    public class FilterList : List<Filter>
    {
        #region Public Methods
        public bool Process(Outlook.MailItem item)
        {
            return this.Any(r => r.Process(item) && r.AbortRuleProcessing);
        }
        #endregion
    }
}
