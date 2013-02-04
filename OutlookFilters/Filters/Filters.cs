using System;
using OutlookFilters.Actions;
using OutlookFilters.Conditions;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookFilters.Filters
{
    public class Filter
    {
        #region Public Properties
        public string Label { get; set; }
        public bool Enabled { get; set; }
        public bool AbortRuleProcessing { get; set; }
        public ConditionExpression Conditions { get; set; }
        public ActionList Actions { get; set; }
        #endregion        

        public Filter()
        {
            Label = String.Empty;
            AbortRuleProcessing = false;
            Conditions = new ConditionExpression();
            Actions = new ActionList();
        }

        #region Public Methods
        /// <summary>
        /// Process the rule with the current MailItem.
        /// </summary>
        /// <param name="item">The Outlook Mail Item</param>
        /// <returns>True if item was processed Successfully; false otherwise.</returns>
        public bool Process(Outlook.MailItem item)
        {
            if (Conditions.Evaluate(item))
            {
                Actions.TrueForAll(a => a.Execute(item)); ;

                return true;
            }

            return false;
        }
        #endregion
    }
}
