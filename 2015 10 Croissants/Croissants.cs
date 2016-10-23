using _2015_10_Croissants.outils;

namespace _2015_10_Croissants
{
    class Croissants
    {
        WindowsHelper windowsHelper;
        CroissantsConfig config;

        public Croissants()
        {
            windowsHelper = new WindowsHelper();
            windowsHelper.ouvrirOutlook();
            config = new CroissantsConfig();
            windowsHelper.sendEmailThroughOutlook(config.sujet, config.contenu, config.destinataires);
        }
    }
}
