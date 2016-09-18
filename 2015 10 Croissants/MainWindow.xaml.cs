using _2015_10_Croissants.outils;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Media;


namespace _2015_10_Croissants {
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window {

        private Timer timer;

        private WindowsHelper windows;
        private CroissantsConfig config;


        public MainWindow() {
            InitializeComponent();

            debug("demarrage");
            debug("ouverture outlook");

            windows = new WindowsHelper(this);

            windows.ouvrirOutlook();
            timer = new Timer();
            timer.Tick += new EventHandler(checkOutlookOpen);
            timer.Interval = 500; // in miliseconds
            timer.Start();

            config = new CroissantsConfig(this);
            configToViews();
        }

        private void checkOutlookOpen(object sender, EventArgs e) {
            if (windows.isOutlookOpen()) {
                vStatutOutlook.Text = "OK";
                vStatutOutlook.Foreground = new SolidColorBrush(Colors.Green);
                vEnvoyer.IsEnabled = true;
                vOuvrirOutlook.Visibility = Visibility.Hidden;
            } else {
                vStatutOutlook.Text = "ferme";
                vStatutOutlook.Foreground = new SolidColorBrush(Colors.Red);
                vEnvoyer.IsEnabled = false;
                vOuvrirOutlook.Visibility = Visibility.Visible;
            }
        }

        private void vOuvrirOutlook_Click(object sender, RoutedEventArgs rea) {
            windows.ouvrirOutlook();
        }

        

        private void vEnvoyer_Click(object sender, RoutedEventArgs rea) {
            debug("vEnvoyer_Click");
            viewsToConfig();
            config.save();
            configToViews();
            debug("envoi mail");
            windows.sendEmailThroughOutlook(config.sujet,config.contenu, config.destinataires);
            debug("fin");
        }

        
        public void debug(string s) {
            s = DateTime.Now + " : " + s + "\n";
            Console.WriteLine(s);
            vDebug.AppendText(s);
            vDebug.ScrollToEnd();
        }

        private void viewsToConfig() {
            config.sujet = vSujet.Text;
            config.destinataires = vDestinataires.Text;
            config.contenu = vContenu.Text;
        }

        private void configToViews() {
            vSujet.Text = config.sujet;
            vDestinataires.Text = config.destinataires;
            vContenu.Text = config.contenu;
        }

        private void vSauvegarder_Click(object sender, RoutedEventArgs e) {
            viewsToConfig();
            config.save();
            configToViews();
        }
    }
}
