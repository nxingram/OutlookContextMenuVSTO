using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
//using Outlook = Microsoft.Office.Interop.Outlook;
//using Office = Microsoft.Office.Core;

//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Xml.Linq;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using stdole;
using System.Drawing;
using System.Windows.Forms;

namespace OutlookContextMenu
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.ItemContextMenuDisplay += ApplicationItemContextMenuDisplay;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Nota: Outlook non genera più questo evento. Se è presente codice che 
            //    deve essere eseguito all'arresto di Outlook, vedere https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region Codice generato da VSTO

        /// <summary>
        /// Metodo richiesto per il supporto della finestra di progettazione. Non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

        // Outlook menu item
        // This triggers whenever a mailitem is rightclicked, and gets a "selection" object passed which contains all selected items
        void ApplicationItemContextMenuDisplay(CommandBar commandBar, Selection selection)
        {
            var cb = commandBar.Controls.Add(MsoControlType.msoControlButton, missing, missing, missing, true) as CommandBarButton;
            if (cb == null) 
                return;
            cb.Visible = true;
            //cb.Picture = ImageConverter.ImageToPictureDisp(Properties.Resources.Desktop.ToBitmap());    // some icon stored in the resources file
            cb.Style = MsoButtonStyle.msoButtonIconAndCaption;                                          // set style to text AND icon
            cb.Click += new _CommandBarButtonEvents_ClickEventHandler(AsterixHook);                     // link click event

            // single MailItem item selection only, NOT 0 based
            if (selection.Count == 1 && selection[1] is MailItem)
            {
                var item = (MailItem)selection[1];                          // retrieve the selected item
                string subject = item.Subject;
                if (subject.Length > 25) subject = subject.Substring(0, 25);// limit max length of the caption
                cb.Caption = "Kakofonix => " + subject;                     // set caption
                cb.Enabled = true;                                          // user selected a single mail item, enable the menu
                cb.Parameter = item.EntryID;                                // this will pass the selected item's identification down when clicked
            }
            else
            {
                cb.Caption = "Kakofonix: Invalid selection";
                cb.Enabled = false;
            }

        }

        // Runs when the actual context menu item is clicked
        private void AsterixHook(CommandBarButton control, ref bool canceldefault)
        {
            string entryid = control.Parameter;                                     // the outlook entry id clicked by the user
            var item = (MailItem)this.Application.Session.GetItemFromID(entryid);   // the actual item
            MessageBox.Show(item.SenderEmailAddress + " " + item.Subject);          // display sender email & subject line

            // further processing
            Console.WriteLine("Pause");
        }
    }


}
