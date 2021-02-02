using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn2
{
    public partial class ThisAddIn
    {


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            CreateCustomFolder(Outlook.OlDefaultFolders.olFolderDrafts, "Draft Smart Email");
            CreateCustomFolder(Outlook.OlDefaultFolders.olFolderSentMail, "Sent Smart Emails");
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

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

        private void CreateCustomFolder(Outlook.OlDefaultFolders parentFolder, string folderName)
        {
            Outlook.Folder folder = (Outlook.Folder)
                Application.ActiveExplorer().Session.GetDefaultFolder
                (parentFolder);
            Outlook.Folder customFolder = null;
            try
            {
                customFolder = (Outlook.Folder)folder.Folders.Add(folderName,
                      Outlook.OlDefaultFolders.olFolderInbox);
                customFolder.WebViewURL = "http://www.google.com";
                customFolder.WebViewOn = true;
                //folder.Folders[folderName].Display();
            }
            catch (Exception ex)
            {
            }
        }

        #endregion
    }
}
