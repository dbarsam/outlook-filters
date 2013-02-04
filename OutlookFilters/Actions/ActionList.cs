using System.Collections.Generic;
using Microsoft.Office.Interop.Outlook;


namespace OutlookFilters.Actions
{
    public class ActionList : List<Action>
    {
        #region Public Methods
        public bool Execute(MailItem item)
        {
            return this.TrueForAll(a => a.Execute(item)); ;
        }
        #endregion
    }
}
