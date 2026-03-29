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

        public string ResolvePath(string reportName, string posId, SaveMode saveMode, string dateLabel = "", string extension = ".pdf")
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

            string fileSuffix = string.IsNullOrWhiteSpace(dateLabel) ? "" : $"_{dateLabel}";
            string baseFileName = $"{reportName}_{posId}{fileSuffix}{extension}";

            if (saveMode == SaveMode.SemcoUpload)
            {
                string directory = Path.Combine(basePath, posId);
                return Path.Combine(directory, baseFileName);
            }
            else if (saveMode == SaveMode.Alternativ)
            {
                return Path.Combine(basePath, baseFileName);
            }
            else
            {
                throw new ArgumentOutOfRangeException(nameof(saveMode), "Unbekannter SaveMode.");
            }
        }
    }
}
