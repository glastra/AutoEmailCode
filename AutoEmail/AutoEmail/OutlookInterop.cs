using System;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace AutoEmail
{
    class OutlookInterop
    {
        Outlook.Application OutlookApp;
        private Outlook.MailItem MailItem;

        public OutlookInterop()
        {
            OutlookApp = new Outlook.Application();
            MailItem = OutlookApp.CreateItemFromTemplate(
              @"c:\GilCats\Autoemail\Templates\default.oft", Type.Missing) as Outlook.MailItem;

        }


        public void CreateMailItem()
        {
            Outlook.MailItem mailItem = (Outlook.MailItem)
                this.OutlookApp.CreateItem(Outlook.OlItemType.olMailItem);


            mailItem.Subject = "This is the subject";
            mailItem.To = "glorialastra76@gmail.com";
            mailItem.Body = "This is the message.";
            mailItem.Importance = Outlook.OlImportance.olImportanceLow;
            mailItem.Display(false);
            mailItem.Send();
        }


        public void CreateItemFromTemplate(string emailTo)
        {

            try
            {
                MailItem.To = emailTo;
                MailItem.Importance = Outlook.OlImportance.olImportanceLow;
                MailItem.Display(false);
                // MailItem.Save();
                MailItem.Send();
                Console.WriteLine("Email sent to : {0}", emailTo);

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while sending email: " + ex);
            }




        }


    }
}
