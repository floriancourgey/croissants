using System;
using System.Collections.Generic;
using System.Diagnostics;

using Outlook = Microsoft.Office.Interop.Outlook;

namespace _2015_10_Croissants.outils {
    public class WindowsHelper {
        
        public void ouvrirOutlook() {
            if (!isOutlookOpen())
                openOutlook();
        }

        public bool isOutlookOpen() {
            string outlook = "outlook";
            Process[] processes = Process.GetProcesses();
            foreach (Process process in processes) {
                try {
                    if (process.ProcessName.ToLower().Equals(outlook) /*||
                        process.MainWindowTitle.ToLower().Contains(outlook)*/
                        ) {
                        Console.WriteLine("outlook deja ouvert (" + process + ")");
                        return true;
                    }
                } catch (Exception e) { Console.WriteLine("Impossible de scanner le process " + process + " (" + e + ")"); }
            }
            Console.WriteLine("outlook non ouvert");
            return false;

        }

        public void openOutlook() {
            Console.WriteLine("ouverture outlook");
            Microsoft.Win32.RegistryKey key =
            Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"Software\microsoft\windows\currentversion\app paths\OUTLOOK.EXE");
            string path = (string)key.GetValue("Path");
            if (path != null) {
                Console.WriteLine("Outlook ouvert");
                Process.Start("OUTLOOK.EXE");
            } else {
                throw new Exception("Outlook introuvable sur ce pc !");
            }
        }

        public void sendEmailThroughOutlook(string sujet, string contenu, string destinataires) {
            try {
                Console.WriteLine("Creation application Outlook");
                Outlook.Application oApp = new Outlook.Application();
                Console.WriteLine("Creation Mail");
                Outlook.MailItem oMsg = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
                Console.WriteLine("Creation contenu");
                oMsg.HTMLBody = contenu.Replace("\n", "<br/>");
                Console.WriteLine("Creation sujet");
                oMsg.Subject = sujet;
                Console.WriteLine("Creation destinataires");
                Outlook.Recipients oRecips = oMsg.Recipients;
                foreach (string destinataire in destinataires.Split(',')) {
                    Console.WriteLine("Ajout " + destinataire);
                    Outlook.Recipient oRecip = oRecips.Add(destinataire);
                    Console.WriteLine("Resolving " + destinataire);
                    oRecip.Resolve();
                    oRecip = null;
                }
                Console.WriteLine("Send.");
                oMsg.Send();
                Console.WriteLine("Clean up.");
                oRecips = null;
                oMsg = null;
                oApp = null;
                Console.WriteLine("Email envoye avec succes.");
            } catch (Exception ex) {
                Console.WriteLine("ERREUR " + ex);
            }
        }
    }
}
