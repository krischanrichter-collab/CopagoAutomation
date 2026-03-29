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
            string reportLabel;

            if (reportName.Contains("ABC", StringComparison.OrdinalIgnoreCase))
            {
                basePath    = saveMode == SaveMode.SemcoUpload ? _settings.AbcBaseFolder : _settings.AbcSammelordnerPath;
                reportLabel = "ABC";
            }
            else if (reportName.Contains("Stundenleistung", StringComparison.OrdinalIgnoreCase))
            {
                basePath    = saveMode == SaveMode.SemcoUpload ? _settings.StundenleistungBaseFolder : _settings.StundenleistungSammelordnerPath;
                reportLabel = "Stundenleistung";
            }
            else
            {
                basePath    = saveMode == SaveMode.SemcoUpload ? _settings.XBaseFolder : _settings.XSammelordnerPath;
                reportLabel = "X-Liste";
            }

            if (string.IsNullOrWhiteSpace(basePath))
            {
                string modeLabel = saveMode == SaveMode.SemcoUpload ? "Semco-Modus (BaseFolder)" : "Alternativ-Modus (Sammelordner)";
                throw new InvalidOperationException($"Pfad für {reportLabel} im {modeLabel} ist nicht konfiguriert.");
            }

            if (saveMode == SaveMode.SemcoUpload)
            {
                string directory = Path.Combine(basePath, posId);
                string fileName  = $"{reportName}_{posId}.pdf";
                return Path.Combine(directory, fileName);
            }
            else if (saveMode == SaveMode.Alternativ)
            {
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
