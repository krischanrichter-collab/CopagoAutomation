using System;

namespace CopagoAutomation.Automation
{
	public readonly struct BoundWindowInfo
	{
		public BoundWindowInfo(IntPtr handle, string title)
		{
			Handle = handle;
			Title = title ?? string.Empty;
		}

		public IntPtr Handle { get; }

		public string Title { get; }

		public bool HasHandle => Handle != IntPtr.Zero;
	}
}