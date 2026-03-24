using System;

namespace CopagoAutomation.Automation
{
		public readonly struct BoundWindowInfo
		{
			public BoundWindowInfo(IntPtr handle, string title, System.Drawing.Rectangle clientRect = default)
			{
				Handle = handle;
				Title = title ?? string.Empty;
				ClientRect = clientRect;
			}

			public IntPtr Handle { get; }

			public string Title { get; }

			public System.Drawing.Rectangle ClientRect { get; }

			public bool HasHandle => Handle != IntPtr.Zero;
		}
}