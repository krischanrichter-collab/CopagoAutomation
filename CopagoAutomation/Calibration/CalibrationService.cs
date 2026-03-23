using System;
using System.Linq;

namespace CopagoAutomation.Calibration
{
	public class CalibrationService
	{
		private readonly CalibrationData _data;

		public CalibrationService(CalibrationData data)
		{
			_data = data ?? throw new ArgumentNullException(nameof(data));
		}

		public CalibrationData GetData()
		{
			return _data;
		}

		public CalibrationModeSet GetOrCreateMode(string modeName)
		{
			ValidateModeName(modeName);

			var mode = _data.Modes.FirstOrDefault(m =>
				string.Equals(m.ModeName, modeName, StringComparison.OrdinalIgnoreCase));

			if (mode != null)
				return mode;

			mode = new CalibrationModeSet
			{
				ModeName = modeName
			};

			_data.Modes.Add(mode);
			return mode;
		}

		public CalibrationProfile GetOrCreateProfile(string modeName, string profileName)
		{
			ValidateModeName(modeName);
			ValidateProfileName(profileName);

			var mode = GetOrCreateMode(modeName);

			var profile = mode.Profiles.FirstOrDefault(p =>
				string.Equals(p.ProfileName, profileName, StringComparison.OrdinalIgnoreCase));

			if (profile != null)
				return profile;

			profile = new CalibrationProfile
			{
				ProfileName = profileName
			};

			mode.Profiles.Add(profile);
			return profile;
		}

		public CalibrationPoint? GetPoint(string modeName, string profileName, string key)
		{
			ValidateModeName(modeName);
			ValidateProfileName(profileName);
			ValidateKey(key);

			var mode = _data.Modes.FirstOrDefault(m =>
				string.Equals(m.ModeName, modeName, StringComparison.OrdinalIgnoreCase));

			if (mode == null)
				return null;

			var profile = mode.Profiles.FirstOrDefault(p =>
				string.Equals(p.ProfileName, profileName, StringComparison.OrdinalIgnoreCase));

			if (profile == null)
				return null;

			return profile.Points.FirstOrDefault(p =>
				string.Equals(p.Key, key, StringComparison.OrdinalIgnoreCase));
		}

		public void SetPoint(string modeName, string profileName, string key, int x, int y)
		{
			ValidateModeName(modeName);
			ValidateProfileName(profileName);
			ValidateKey(key);
			ValidateCoordinate(x, nameof(x));
			ValidateCoordinate(y, nameof(y));

			var profile = GetOrCreateProfile(modeName, profileName);

			var existingPoint = profile.Points.FirstOrDefault(p =>
				string.Equals(p.Key, key, StringComparison.OrdinalIgnoreCase));

			if (existingPoint != null)
			{
				existingPoint.X = x;
				existingPoint.Y = y;
				return;
			}

			profile.Points.Add(new CalibrationPoint
			{
				Key = key,
				X = x,
				Y = y
			});
		}

		public bool HasPoint(string modeName, string profileName, string key)
		{
			return GetPoint(modeName, profileName, key) != null;
		}

		public bool IsProfileComplete(string modeName, string profileName)
		{
			ValidateModeName(modeName);
			ValidateProfileName(profileName);

			var requiredKeys = CalibrationDefinitions.GetKeysForProfile(profileName);
			if (requiredKeys.Count == 0)
				return false;

			var mode = _data.Modes.FirstOrDefault(m =>
				string.Equals(m.ModeName, modeName, StringComparison.OrdinalIgnoreCase));

			if (mode == null)
				return false;

			var profile = mode.Profiles.FirstOrDefault(p =>
				string.Equals(p.ProfileName, profileName, StringComparison.OrdinalIgnoreCase));

			if (profile == null)
				return false;

			return requiredKeys.All(requiredKey =>
				profile.Points.Any(point =>
					string.Equals(point.Key, requiredKey, StringComparison.OrdinalIgnoreCase)));
		}

		private static void ValidateModeName(string modeName)
		{
			if (string.IsNullOrWhiteSpace(modeName))
				throw new ArgumentException("modeName darf nicht leer sein.", nameof(modeName));
		}

		private static void ValidateProfileName(string profileName)
		{
			if (string.IsNullOrWhiteSpace(profileName))
				throw new ArgumentException("profileName darf nicht leer sein.", nameof(profileName));
		}

		private static void ValidateKey(string key)
		{
			if (string.IsNullOrWhiteSpace(key))
				throw new ArgumentException("key darf nicht leer sein.", nameof(key));
		}

		private static void ValidateCoordinate(int value, string parameterName)
		{
			if (value < 0)
				throw new ArgumentOutOfRangeException(parameterName, "Koordinaten dürfen nicht negativ sein.");
		}
	}
}