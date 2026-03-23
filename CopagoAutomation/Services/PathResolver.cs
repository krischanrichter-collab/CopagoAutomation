using System;
using System.IO;
using CopagoAutomation.Models;

namespace CopagoAutomation.Services
{
    public class PathResolver
    {
        private readonly AppSettings _settings;

        public PathResolver(AppSettings settings)
        {
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
        }

        public string ResolvePath(string reportName, string posId, SaveMode saveMode)
        {
            string basePath;

            if (saveMode == SaveMode.SemcoUpload)
            {
                // Im Semco-Modus wird der BaseFolder aus den Einstellungen verwendet
                // und der Pfad ist fest definiert (BaseFolder\POS\ReportName.pdf)
                if (string.IsNullOrWhiteSpace(_settings.BaseFolder))
                {
                    throw new InvalidOperationException("BaseFolder ist im Semco-Modus nicht konfiguriert.");
                }
                basePath = _settings.BaseFolder;
            }
            else if (saveMode == SaveMode.Alternativ)
            {
                // Im Alternativ-Modus wird der SammelordnerPath aus den Einstellungen verwendet
                if (string.IsNullOrWhiteSpace(_settings.SammelordnerPath))
                {
                    throw new InvalidOperationException("Alternativer Ordner ist im Alternativ-Modus nicht konfiguriert.");
                }
                basePath = _settings.SammelordnerPath;
            }
            else
            {
                throw new ArgumentOutOfRangeException(nameof(saveMode), "Unbekannter SaveMode.");
            }

            // Sicherstellen, dass der Basispfad mit einem Verzeichnistrennzeichen endet
            if (!basePath.EndsWith(Path.DirectorySeparatorChar.ToString()))
            {
                basePath += Path.DirectorySeparatorChar;
            }

            // Dateiname basierend auf Report und POS
            // Beispiel: ABC_Analyse_POS123.pdf
            string fileName = $"{reportName}_{posId}.pdf";

            // Finaler Pfad
            return Path.Combine(basePath, fileName);
        }
    }
}
