using System;
using System.Collections.Generic;
using System.Diagnostics;

using Outlook = Microsoft.Office.Interop.Outlook;

namespace _2015_10_Croissants.outils {
    public class WindowsHelper {

        private MainWindow mainWindow;

        public WindowsHelper(MainWindow mainWindow) {
            this.mainWindow = mainWindow;
        }

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
                } catch (Exception e) { mainWindow.debug("Impossible de scanner le process " + process + " (" + e + ")"); }
            }
            Console.WriteLine("outlook non ouvert");
            return false;

        }

        public void openOutlook() {
            mainWindow.debug("ouverture outlook");
            Microsoft.Win32.RegistryKey key =
            Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"Software\microsoft\windows\currentversion\app paths\OUTLOOK.EXE");
            string path = (string)key.GetValue("Path");
            if (path != null) {
                mainWindow.debug("Outlook ouvert");
                Process.Start("OUTLOOK.EXE");
            } else {
                throw new Exception("Outlook introuvable sur ce pc !");
            }
        }

        public void sendEmailThroughOutlook(string sujet, string contenu, string destinataires) {
            try {
                mainWindow.debug("Creation application Outlook");
                Outlook.Application oApp = new Outlook.Application();
                mainWindow.debug("Creation Mail");
                Outlook.MailItem oMsg = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
                mainWindow.debug("Creation contenu");
                oMsg.HTMLBody = contenu.Replace("\n", "<br/>");
                mainWindow.debug("Creation sujet");
                oMsg.Subject = sujet;
                mainWindow.debug("Creation destinataires");
                Outlook.Recipients oRecips = oMsg.Recipients;
                foreach (string destinataire in destinataires.Split(',')) {
                    mainWindow.debug("Ajout " + destinataire);
                    Outlook.Recipient oRecip = oRecips.Add(destinataire);
                    mainWindow.debug("Resolving " + destinataire);
                    oRecip.Resolve();
                    oRecip = null;
                }
                mainWindow.debug("Send.");
                oMsg.Send();
                mainWindow.debug("Clean up.");
                oRecips = null;
                oMsg = null;
                oApp = null;
                mainWindow.debug("Email envoye avec succes.");
            } catch (Exception ex) {
                mainWindow.debug("ERREUR " + ex);
            }
        }
    }
}
