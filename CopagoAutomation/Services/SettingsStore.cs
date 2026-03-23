using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using CopagoAutomation.Models;

namespace CopagoAutomation.Services
{
	public sealed class SettingsStore
	{
		private readonly string _path;
		private readonly JsonSerializerOptions _json;

		public SettingsStore(string path)
		{
			_path = path;
			_json = new JsonSerializerOptions
			{
				WriteIndented = true,
				AllowTrailingCommas = true,
				ReadCommentHandling = JsonCommentHandling.Skip,
				DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
			};
		}

		public async Task<AppSettings> LoadAsync()
		{
			if (!File.Exists(_path))
				return new AppSettings();

			var json = await File.ReadAllTextAsync(_path).ConfigureAwait(false);
			var settings = JsonSerializer.Deserialize<AppSettings>(json, _json);
			return settings ?? new AppSettings();
		}

		public async Task SaveAsync(AppSettings settings)
		{
			Directory.CreateDirectory(Path.GetDirectoryName(_path) ?? ".");
			var json = JsonSerializer.Serialize(settings, _json);
			await File.WriteAllTextAsync(_path, json).ConfigureAwait(false);
		}
	}
}