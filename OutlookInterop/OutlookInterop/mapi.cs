using System.Runtime.InteropServices;
using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Reflection;
using Microsoft.Office.Interop.Outlook;

namespace OutlookInterop
{
    public static class Mapi
    {
        [UnmanagedCallersOnly]
        public static long GetMailReceiveDate(IntPtr mailEntryIdPtr, int textLength)
        {
            try
            {
                var mailEntryId = Marshal.PtrToStringUTF8(mailEntryIdPtr, textLength);

                // Create the Outlook application.
                // in-line initialization
                Outlook.Application oApp = new Outlook.Application();

                // Get the MAPI namespace.
                Outlook.NameSpace oNS = oApp.GetNamespace("mapi");

                // Log on by using the default profile or existing session (no dialog box).
                oNS.Logon(Missing.Value, Missing.Value, false, true);

                var mailEntry = oNS.GetItemFromID(mailEntryId) as MailItem;
                // Console.WriteLine(mailEntry.ReceivedTime);
                // mailEntry.Display();

                oNS.Logoff();
                //Explicitly release objects.
                oNS = null;
                oApp = null;

                return ((DateTimeOffset)mailEntry.ReceivedTime.ToUniversalTime()).ToUnixTimeSeconds();
            }
            //Error handler.
            catch (System.Exception e)
            {
                // Console.WriteLine("{0} Exception caught: ", e);
            }

            return 0;
        }
    }
}