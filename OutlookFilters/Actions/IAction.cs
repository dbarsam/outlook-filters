using Microsoft.Office.Interop.Outlook;

namespace OutlookFilters.Actions
{
    public interface IAction
    {
        bool Execute(MailItem item);
    }
}