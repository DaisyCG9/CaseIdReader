﻿using Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace dotNetEmailReader
{
    public class OutlookEmails
    {
        public string EmailFrom { get; set; }
        public string EmailSubject { get; set; }
        public string EmailBody { get; set; }
        public string EmailTo { get; set; }
        public DateTime EmailDate { get; set; }


        public static List<OutlookEmails> ReadMailItems()
        { 
            Application outlookApplication = null;
            NameSpace outlookNamespace = null;
            MAPIFolder inboxFolder = null;

            Items mailItems = null;
            List<OutlookEmails> listEmailDetails = new List<OutlookEmails>();
            OutlookEmails  emailDetails;
            try
            {
                outlookApplication = new Application();
                outlookNamespace = outlookApplication.GetNamespace("MAPI");
                inboxFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                mailItems = inboxFolder.Items;
                foreach (dynamic item in mailItems)
               {
                    if(item is MailItem)
                    {
                           emailDetails = new OutlookEmails();
                           emailDetails.EmailFrom = item.SenderEmailAddress;
                           emailDetails.EmailSubject = item.Subject;
                           emailDetails.EmailBody = item.Body;
                           emailDetails.EmailTo = item.To;
                           emailDetails.EmailDate = item.ReceivedTime;
                        
                           listEmailDetails.Add(emailDetails);
                       
                           ReleaseComObject(item);
                    }
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                ReleaseComObject(mailItems);
                ReleaseComObject(inboxFolder);
                ReleaseComObject(outlookNamespace);
                ReleaseComObject(outlookApplication);
            }
            return listEmailDetails;
        }
        private static void ReleaseComObject(object obj)
        {
            if(obj !=null)
            {
                Marshal.ReleaseComObject(obj);
                obj = null;
            }
        }
    }
}
