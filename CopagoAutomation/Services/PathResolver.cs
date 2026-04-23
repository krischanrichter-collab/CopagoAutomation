using System;
using System.IO;
using System.Linq;
using CopagoAutomation.Models;

namespace CopagoAutomation.Services
{
    public class PathResolver
    {
        private const string SemcoStandortePath = @"F:\Gebietsverkaufsleiter\Semco\Standorte";

        private readonly AppSettings _settings;

        public PathResolver(AppSettings settings)
        {
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
        }

        public string ResolvePath(string reportName, string posId, SaveMode saveMode, string dateLabel = "", string extension = ".pdf")
        {
            string fileSuffix   = string.IsNullOrWhiteSpace(dateLabel) ? "" : $"_{dateLabel}";
            string baseFileName = $"{reportName}_{posId}{fileSuffix}{extension}";

            if (saveMode == SaveMode.SemcoUpload)
                return ResolveSemcoPath(reportName, posId, baseFileName);

            if (saveMode == SaveMode.Alternativ)
                return ResolveAlternativPath(reportName, baseFileName);

            throw new ArgumentOutOfRangeException(nameof(saveMode), "Unbekannter SaveMode.");
        }

        private string ResolveSemcoPath(string reportName, string posId, string baseFileName)
        {
            if (reportName.Contains("ABC", StringComparison.OrdinalIgnoreCase))
            {
                string posFolder = FindSemcoPosFolder(posId);
                return Path.Combine(posFolder, "ABC Statistiken", baseFileName);
            }

            if (reportName.Contains("X-Liste", StringComparison.OrdinalIgnoreCase))
            {
                string posFolder = FindSemcoPosFolder(posId);
                return Path.Combine(posFolder, "X-Listen", baseFileName);
            }

            // Stundenleistung, Artikelfrequenz: konfigurierter BaseFolder
            string? basePath = GetBaseFolder(reportName);
            if (string.IsNullOrWhiteSpace(basePath))
                throw new InvalidOperationException($"Semco-Pfad für '{reportName}' ist nicht konfiguriert.");

            return Path.Combine(basePath, posId, baseFileName);
        }

        private string ResolveAlternativPath(string reportName, string baseFileName)
        {
            string? sammelordner = GetSammelordnerPath(reportName);
            if (string.IsNullOrWhiteSpace(sammelordner))
                throw new InvalidOperationException($"Sammelordner für '{reportName}' ist nicht konfiguriert.");

            return Path.Combine(sammelordner, baseFileName);
        }

        private static string FindSemcoPosFolder(string posId)
        {
            if (!Directory.Exists(SemcoStandortePath))
                throw new InvalidOperationException($"Semco-Standorte-Ordner nicht gefunden: {SemcoStandortePath}");

            // Erst mit führenden Nullen auf 3 Stellen suchen (z.B. "17" → "017*"),
            // dann als Fallback die Nummer wie eingegeben.
            string paddedId = posId.PadLeft(3, '0');
            string? match = Directory.GetDirectories(SemcoStandortePath, paddedId + "*").FirstOrDefault()
                         ?? Directory.GetDirectories(SemcoStandortePath, posId + "*").FirstOrDefault();

            if (match == null)
                throw new InvalidOperationException($"Kein Semco-Ordner für POS '{posId}' in {SemcoStandortePath} gefunden.");

            return match;
        }

        private string? GetBaseFolder(string reportName)
        {
            if (reportName.Contains("Stundenleistung", StringComparison.OrdinalIgnoreCase))
                return _settings.StundenleistungBaseFolder;
            if (reportName.Contains("Artikelfrequenz", StringComparison.OrdinalIgnoreCase))
                return _settings.ArtikelfrequenzBaseFolder;
            return null;
        }

        private string? GetSammelordnerPath(string reportName)
        {
            if (reportName.Contains("ABC", StringComparison.OrdinalIgnoreCase))
                return _settings.AbcSammelordnerPath;
            if (reportName.Contains("X-Liste", StringComparison.OrdinalIgnoreCase))
                return _settings.XSammelordnerPath;
            if (reportName.Contains("Stundenleistung", StringComparison.OrdinalIgnoreCase))
                return _settings.StundenleistungSammelordnerPath;
            if (reportName.Contains("Artikelfrequenz", StringComparison.OrdinalIgnoreCase))
                return _settings.ArtikelfrequenzSammelordnerPath;
            return null;
        }
    }
}
