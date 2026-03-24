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
            string? basePath;

            // Wir müssen unterscheiden, ob es ein ABC- oder X-Liste-Report ist, 
            // um die richtigen Einstellungen zu laden.
            bool isAbc = reportName.Contains("ABC", StringComparison.OrdinalIgnoreCase);

            if (saveMode == SaveMode.SemcoUpload)
            {
                // Im Semco-Modus wird der BaseFolder aus den Einstellungen verwendet
                // und der Pfad ist fest definiert (BaseFolder\POS\ReportName.pdf)
                basePath = isAbc ? _settings.AbcBaseFolder : _settings.XBaseFolder;
                
                if (string.IsNullOrWhiteSpace(basePath))
                {
                    throw new InvalidOperationException($"BaseFolder für {(isAbc ? "ABC" : "X-Liste")} ist im Semco-Modus nicht konfiguriert.");
                }

                // Für Semco-Upload wird der Pfad automatisch generiert: BaseFolder\POS-ID\ReportName_POS-ID.pdf
                // Beispiel: C:\CopagoReports\POS123\ABC_Analyse_POS123.pdf
                string directory = Path.Combine(basePath, posId);
                string fileName = $"{reportName}_{posId}.pdf";
                return Path.Combine(directory, fileName);
            }
            else if (saveMode == SaveMode.Alternativ)
            {
                // Im Alternativ-Modus wird der SammelordnerPath aus den Einstellungen verwendet
                // und der Pfad ist fest definiert (SammelordnerPath\ReportName_POS-ID.pdf)
                basePath = isAbc ? _settings.AbcSammelordnerPath : _settings.XSammelordnerPath;

                if (string.IsNullOrWhiteSpace(basePath))
                {
                    throw new InvalidOperationException($"Alternativer Ordner für {(isAbc ? "ABC" : "X-Liste")} ist im Alternativ-Modus nicht konfiguriert.");
                }

                // Für Alternativ-Modus wird der Pfad automatisch generiert: SammelordnerPath\ReportName_POS-ID.pdf
                // Beispiel: C:\AlternativReports\ABC_Analyse_POS123.pdf
                string fileName = $"{reportName}_{posId}.pdf";
                return Path.Combine(basePath, fileName);
            }
            else
            {
                throw new ArgumentOutOfRangeException(nameof(saveMode), "Unbekannter SaveMode.");
            }
        }
    }
}
