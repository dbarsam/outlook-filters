using OutlookFilters.Ribbons;

namespace OutlookFilters.Ribbons
{
    partial class RibbonUI : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonUI()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tabOutlookFilters = this.Factory.CreateRibbonTab();
            this.tabOutlookFilters.SuspendLayout();
            // 
            // tabOutlookFilters
            // 
            this.tabOutlookFilters.Label = "Outlook Filters";
            this.tabOutlookFilters.Name = "tabOutlookFilters";
            // 
            // RibbonUI
            // 
            this.Name = "RibbonUI";
            this.RibbonType = "Microsoft.Outlook.Mail.Read";
            this.Tabs.Add(this.tabOutlookFilters);
            this.tabOutlookFilters.ResumeLayout(false);
            this.tabOutlookFilters.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabOutlookFilters;
    }

}
namespace OutlookFilters
{
    partial class ThisRibbonCollection
    {
        internal RibbonUI Ribbon
        {
            get { return this.GetRibbon<RibbonUI>(); }
        }
    }
}
