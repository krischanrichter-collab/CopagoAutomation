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

            // Wir müssen unterscheiden, ob es ein ABC- oder X-Liste-Report ist, 
            // um die richtigen Einstellungen zu laden.
            bool isAbc = reportName.Contains("ABC", StringComparison.OrdinalIgnoreCase);

            if (saveMode == SaveMode.SemcoUpload)
            {
                basePath = isAbc ? _settings.AbcBaseFolder : _settings.XBaseFolder;
                
                if (string.IsNullOrWhiteSpace(basePath))
                {
                    throw new InvalidOperationException($"BaseFolder für {(isAbc ? "ABC" : "X-Liste")} ist im Semco-Modus nicht konfiguriert.");
                }
            }
            else if (saveMode == SaveMode.Alternativ)
            {
                basePath = isAbc ? _settings.AbcSammelordnerPath : _settings.XSammelordnerPath;

                if (string.IsNullOrWhiteSpace(basePath))
                {
                    throw new InvalidOperationException($"Alternativer Ordner für {(isAbc ? "ABC" : "X-Liste")} ist im Alternativ-Modus nicht konfiguriert.");
                }
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
