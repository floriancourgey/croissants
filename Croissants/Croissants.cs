using _2015_10_Croissants.outils;

namespace _2015_10_Croissants
{
    class Croissants
    {
        public static void Main(string[] args)
        {
            WindowsHelper windowsHelper = new WindowsHelper();
            windowsHelper.ouvrirOutlook();
            CroissantsConfig config = new CroissantsConfig();
            windowsHelper.sendEmailThroughOutlook(config.sujet, config.contenu, config.destinataires);
        }
    }
}
