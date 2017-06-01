

using System.Drawing;

namespace fncGTD4Outlook.Comun
{
    public static class Constants
    {
        public static Color THEME_DARK_PANEL_BACKCOLOR = Color.FromArgb(106, 106, 106);
        public static Color THEME_DARK_PANEL_FORECOLOR = Color.FromArgb(255, 255, 255);
        public static Color THEME_LIGHT_PANEL_BACKCOLOR = Color.FromArgb(255, 255, 255);
        public static Color THEME_LIGHT_PANEL_FORECOLOR = Color.FromArgb(106, 106, 106);

        //TODO: incluir en un archivo de configuracion
        public static string folderArchivar = "1-Archivar";
        public static string folderDelegar = "2-Waiting";
        public static string folderDiferir = "3-Deferred";
        public static string folderConservar = "4-Someday";
        public static string folderReferencia = "5-Reference";
        public static string folderRecurrente = "6-Recurrence";

        public static char emailDelimiter = ';';

        public static string myEmail = "fernando.novoa@pe.engie.com";
    }
}
