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
            this.groupEdit = this.Factory.CreateRibbonGroup();
            this.buttonCreateFilter = this.Factory.CreateRibbonButton();
            this.buttonManageFilters = this.Factory.CreateRibbonButton();
            this.tabOutlookFilters.SuspendLayout();
            this.groupEdit.SuspendLayout();
            // 
            // tabOutlookFilters
            // 
            this.tabOutlookFilters.Groups.Add(this.groupEdit);
            this.tabOutlookFilters.Label = "Outlook Filters";
            this.tabOutlookFilters.Name = "tabOutlookFilters";
            // 
            // groupEdit
            // 
            this.groupEdit.Items.Add(this.buttonCreateFilter);
            this.groupEdit.Items.Add(this.buttonManageFilters);
            this.groupEdit.Label = "Edit Filters";
            this.groupEdit.Name = "groupEdit";
            // 
            // buttonCreateFilter
            // 
            this.buttonCreateFilter.ImageName = "FilterNew";
            this.buttonCreateFilter.Label = "Create Filter...";
            this.buttonCreateFilter.Name = "buttonCreateFilter";
            this.buttonCreateFilter.OfficeImageId = "CreateMailRule";
            this.buttonCreateFilter.ShowImage = true;
            this.buttonCreateFilter.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonCreateFilter_Click);
            // 
            // buttonManageFilters
            // 
            this.buttonManageFilters.ImageName = "AdvancedFilterDialog";
            this.buttonManageFilters.Label = "Manager Filters...";
            this.buttonManageFilters.Name = "buttonManageFilters";
            this.buttonManageFilters.ShowImage = true;
            this.buttonManageFilters.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonManageFilters_Click);
            // 
            // RibbonUI
            // 
            this.Name = "RibbonUI";
            this.RibbonType = "Microsoft.Outlook.Mail.Read";
            this.Tabs.Add(this.tabOutlookFilters);
            this.tabOutlookFilters.ResumeLayout(false);
            this.tabOutlookFilters.PerformLayout();
            this.groupEdit.ResumeLayout(false);
            this.groupEdit.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabOutlookFilters;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupEdit;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCreateFilter;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonManageFilters;
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
