using System;
using System.Collections.Generic;

namespace CopagoAutomation.Models
{
	public enum MachineMode
	{
		Laptop = 0,
		Docking = 1
	}

	// NEU: Speicher-Modus
	public enum SaveMode
	{
		SemcoUpload = 0,
		Alternativ = 1
	}

	public sealed class AppSettings
	{
		public int Version { get; set; } = 1;

		// Maschinenmodus pro Report
		public MachineMode AbcMode { get; set; } = MachineMode.Laptop;
		public MachineMode XMode { get; set; } = MachineMode.Laptop;

		// Letzter POS (oder null wenn keiner)
		public string? LastPosId { get; set; }

		// Multi-Auswahl gespeicherter POS
		public List<string> SelectedPosIds { get; set; } = new();

		// =========================================
		// 🔥 NEUE SPEICHER-ARCHITEKTUR
		// =========================================

		// Welcher Modus ist aktiv
		public SaveMode AbcSaveMode { get; set; } = SaveMode.SemcoUpload;
public SaveMode XSaveMode { get; set; } = SaveMode.SemcoUpload;

		// Basisordner für Semco-Modus
		// Beispiel: C:\CopagoReports\
		public string? AbcBaseFolder { get; set; }
public string? XBaseFolder { get; set; }

		// Wird im Alternativ-Modus verwendet
		// → entspricht dem früheren Sammelordner
		public string? AbcSammelordnerPath { get; set; }
public string? XSammelordnerPath { get; set; }

		// =========================================
		// (ALT) TEMP / Sammel-Logik (vorerst behalten)
		// =========================================
		public bool UseTempPdfStrategy { get; set; } = true;

		// =========================================
		// Kalibrierung
		// =========================================
		public CalibrationProfile LaptopCalibration { get; set; } = CalibrationProfile.CreateDefault();
		public CalibrationProfile DockingCalibration { get; set; } = CalibrationProfile.CreateDefault();

		// =========================================
		// POS Mapping (für späteren PathResolver)
		// =========================================
		public Dictionary<string, PosDefinition> PosMap { get; set; } = new();
	}

	public sealed class PosDefinition
	{
		public string PosId { get; set; } = "";
		public string DisplayName { get; set; } = "";

		// Wird später für Semco-Pfade genutzt
		public string? LocalBasePath { get; set; }

		// Für Upload / FTP
		public string? RemoteBasePath { get; set; }
	}
}