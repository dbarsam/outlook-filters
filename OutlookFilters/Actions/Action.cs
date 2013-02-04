using System.Xml.Serialization;
using Microsoft.Office.Interop.Outlook;

namespace OutlookFilters.Actions
{
    public abstract class Action : IAction
    {
        public abstract bool Execute(MailItem item);
    }
}