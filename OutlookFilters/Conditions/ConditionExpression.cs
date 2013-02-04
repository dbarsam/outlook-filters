using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Outlook;

namespace OutlookFilters.Conditions
{
    public class ConditionExpression : Condition
    {
        #region Enum
        public enum LogicOperatorType
        {
            Identity,
            Not,
            And,
            Or,
            Xor
        }
        #endregion

        #region Properties
        public LogicOperatorType Operator { get; set; }
        public List<Condition> Conditions { get; set; }
        #endregion

        #region Constructor
        public ConditionExpression()
        {
            Conditions = new List<Condition>();
        }
        #endregion

        #region Condition Implementation
        public override bool Evaluate(MailItem item)
        {
            if (item == null)
                return false;

            if (!Conditions.Any())
                return false;

            try
            {
                switch (Operator)
                {
                    case LogicOperatorType.Identity:
                        return Conditions.FirstOrDefault().Evaluate(item);
                    case LogicOperatorType.Not:
                        return Conditions.FirstOrDefault().Evaluate(item);
                    case LogicOperatorType.And:
                        return Conditions.All(c => c.Evaluate(item));
                    case LogicOperatorType.Or:
                        return Conditions.Any(c => c.Evaluate(item));
                    case LogicOperatorType.Xor:
                        return Conditions.Where(c => c.Evaluate(item)).Take(2).Count() == 1;
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
