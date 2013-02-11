using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Xml.Serialization;
using OutlookFilters.Filters;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace OutlookFilters
{
    public partial class ThisAddIn
    {
        private FilterList _Filters = new FilterList();
    
        private Outlook.Explorers _Explorers;
        private Outlook.Inspectors _Inspectors;
        private Outlook.NameSpace _Namespace;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _Explorers = this.Application.Explorers;
            _Inspectors = this.Application.Inspectors;
            _Namespace = this.Application.GetNamespace("MAPI");
            
            _Explorers.Application.NewMailEx += new Outlook.ApplicationEvents_11_NewMailExEventHandler(Application_NewMailEx);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void Application_NewMailEx(string EntryID)
        {
            var newMail = (Outlook.MailItem)_Explorers.Application.Session.GetItemFromID(EntryID, System.Reflection.Missing.Value);

            if (newMail != null)
            {
                _Filters.Any(r => r.Process(newMail) && r.CanAbortProcessing);
            }
        }

        #region Filter Definition File Serialization
        private bool SaveFilters(string path)
        {
            var serializer = new XmlSerializer(_Filters.GetType());
            using (TextWriter textWriter = new StreamWriter(path))
            {
                serializer.Serialize(textWriter, _Filters);
            }

            return true;
        }

        private bool LoadFilters(string path)
        {
            if (!File.Exists(path))
                return false;

            var serializer = new XmlSerializer(_Filters.GetType());
            using (TextReader textReader = new StreamReader(path))
            {
                _Filters = (FilterList)serializer.Deserialize(textReader);
            }
            return true;
        }
        #endregion

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion

        #region Test Methods
        private void TestSerialization()
        {
            var rule = new Filter() { Enabled = true, CanAbortProcessing = false, Label = "Test" };
            rule.Actions.Add(new OutlookFilters.Actions.MoveAction() { DestinationFolder = this.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox) });
            rule.Conditions.Conditions.Add(new OutlookFilters.Conditions.TextMatch() { SearchText = "Text", MatchMethod = Conditions.TextMatch.TextMatchType.Exact, FieldType = Conditions.TextMatch.TextFieldType.Subject });

            _Filters.Add(rule);

            SaveFilters("C:\\filterlist.xml");
            LoadFilters("C:\\filterlist.xml");
        }
        #endregion

        #region Ribbon Implementation
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new OutlookFilters.Ribbons.Ribbon();
        }
        #endregion
    }
}
