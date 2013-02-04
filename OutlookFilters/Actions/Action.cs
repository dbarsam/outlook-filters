using System.Xml.Serialization;
using Microsoft.Office.Interop.Outlook;

namespace OutlookFilters.Actions
{
    [XmlInclude(typeof(MoveAction))]
    public abstract class Action : IAction
    {
        public abstract bool Execute(MailItem item);
    }
}