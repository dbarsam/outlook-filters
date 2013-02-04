using Microsoft.Office.Interop.Outlook;

namespace OutlookFilters.Conditions
{
    public interface ICondition
    {
        bool Evaluate(MailItem item);
    }
}
