using System.Xml.Serialization;
using Microsoft.Office.Interop.Outlook;

namespace OutlookFilters.Conditions
{
    [
    XmlInclude(typeof(ConditionExpression))
    ]
    public abstract class Condition : ICondition
    {
        public abstract bool Evaluate(MailItem item);
    }
}