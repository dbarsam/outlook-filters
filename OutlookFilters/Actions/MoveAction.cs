using System;
using System.Xml.Serialization;
using Microsoft.Office.Interop.Outlook;


namespace OutlookFilters.Actions
{
    public class MoveAction : Action
    {
        #region Properties
        public string DestinationPath 
        {
            get { return DestinationFolder != null ? DestinationFolder.FolderPath : String.Empty; }
            set { DestinationFolder = FindFolder(value); }
        }

        [XmlIgnore]
        public MAPIFolder DestinationFolder { get; set; }
        #endregion

        #region IAction Implementation
        public override bool Execute(MailItem item)
        {         
            MAPIFolder folder = null;
            item.Move(DestinationFolder);

            return true;
        }
        #endregion

        #region Protected Methods
        protected MAPIFolder FindFolder(string path)
        {
            return null;            
        }
        #endregion
    }
}
