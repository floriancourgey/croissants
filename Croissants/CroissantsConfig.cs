using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace _2015_10_Croissants {
    public class CroissantsConfig {

        public string sujet {
            get { return doc.SelectSingleNode("/" + LOG_NAME + "/sujet").InnerText; }
            set { doc.SelectSingleNode("/" + LOG_NAME + "/sujet").InnerText = value; }
        }
        public string contenu {
            get { return doc.SelectSingleNode("/" + LOG_NAME + "/contenu").InnerText; }
            set { doc.SelectSingleNode("/" + LOG_NAME + "/contenu").InnerText = value; }
        }
        public string destinataires {
            get { return doc.SelectSingleNode("/" + LOG_NAME + "/destinataires").InnerText; }
            set { doc.SelectSingleNode("/" + LOG_NAME + "/destinataires").InnerText = value; }
        }

        private static readonly object Locker = new object();
        private static XmlDocument doc = null;
        private const string LOG_NAME = "configuration";
        private string LOG_PATH = @"" + LOG_NAME+".xml";

        public void save() {
            Console.WriteLine("saving...");
            doc.Save(LOG_PATH);
            Console.WriteLine("save OK");
        }

        public CroissantsConfig() {
            Console.WriteLine("ouverture config "+ LOG_PATH);
            doc = new XmlDocument();
            if (File.Exists(LOG_PATH)) {
                Console.WriteLine("config deja existante");
                doc.Load(LOG_PATH);
            } else {
                Console.WriteLine("config non existante, creation");
                XmlElement root = doc.CreateElement(LOG_NAME);
                root.AppendChild(doc.CreateElement("sujet"));
                root.AppendChild(doc.CreateElement("contenu"));
                root.AppendChild(doc.CreateElement("destinataires"));
                doc.AppendChild(root);
                doc.Save(LOG_PATH);

                sujet = "Je paye mes croissants le XYZ";
                contenu = "Conte-nu.\nCeci est le contenu youhou";
                destinataires = "oo@oo.com , ab@cd.ef";

                
                doc.Save(LOG_PATH);
                Console.WriteLine("config creee");
            }
        }

        public List<string> xmlTodestinataires(string sDestinataires) {
            List<String> destinataires = new List<string>();
            sDestinataires = sDestinataires.Replace(" ", "");
            foreach(string destinataire in sDestinataires.Split(',')) {
                destinataires.Add(destinataire);
            }
            return destinataires;
        }
        public string destinatairesToXml(List <string> destinataires) {
            return string.Join<string>(" , ", destinataires);
        }

        private XmlElement ExceptionToxmlElement(Exception e, int stackNumber) {
            XmlElement exception = doc.CreateElement("exception_" + stackNumber);
            exception.AppendChild(doc.CreateElement("message")).InnerText = e.Message;
            exception.AppendChild(doc.CreateElement("source")).InnerText = e.Source;
            if (e.StackTrace != null) {
                XmlElement stackTrace = (XmlElement)exception.AppendChild(doc.CreateElement("stackTrace_" + stackNumber));
                string[] callers = e.StackTrace.Split('\n');
                foreach (string caller in callers) {
                    stackTrace.AppendChild(doc.CreateElement("caller")).InnerText = caller.Trim('\r', '\n', ' ');
                }
            }
            //.InnerText = e.StackTrace;
            return exception;
        }
    }
}
