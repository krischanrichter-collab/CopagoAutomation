using System;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;

namespace CopagoAutomation.Calibration
{
	public class CalibrationStorage
	{
		private readonly string _filePath;

		private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
		{
			WriteIndented = true
		};

		public CalibrationStorage(string filePath)
		{
			if (string.IsNullOrWhiteSpace(filePath))
				throw new ArgumentException("filePath darf nicht leer sein.", nameof(filePath));

			_filePath = filePath;
		}

		public async Task<CalibrationData> LoadAsync()
		{
			if (!File.Exists(_filePath))
				return new CalibrationData();

			string json = await File.ReadAllTextAsync(_filePath);

			if (string.IsNullOrWhiteSpace(json))
				return new CalibrationData();

			var data = JsonSerializer.Deserialize<CalibrationData>(json, JsonOptions) ?? new CalibrationData();

			data.Modes ??= new();

			foreach (var mode in data.Modes)
			{
				mode.ModeName ??= string.Empty;
				mode.Profiles ??= new();

				foreach (var profile in mode.Profiles)
				{
					profile.ProfileName ??= string.Empty;
					profile.Points ??= new();

					foreach (var point in profile.Points)
					{
						point.Key ??= string.Empty;
					}
				}
			}

			return data;
		}

		public async Task SaveAsync(CalibrationData data)
		{
			if (data == null)
				throw new ArgumentNullException(nameof(data));

			string? directory = Path.GetDirectoryName(_filePath);
			if (!string.IsNullOrWhiteSpace(directory))
				Directory.CreateDirectory(directory);

			string json = JsonSerializer.Serialize(data, JsonOptions);
			await File.WriteAllTextAsync(_filePath, json);
		}
	}
}