
using System.Runtime.InteropServices;
using System.Text;

namespace ExcelFusion
{
    internal partial class NativeMethods
    {
        /// <summary>
        /// Gets the handle of the message box with the specified title.
        /// </summary>
        /// <param name="title">The title of the message box to get the handle of.</param>
        /// <returns>The handle of the message box with the specified title.</returns>
        public static IntPtr GetMessageBoxHandle(string title)
        {
            IntPtr hWnd = IntPtr.Zero;
            while ((hWnd = GetWindow(hWnd, 5)) != IntPtr.Zero)
            {
                StringBuilder sb = new StringBuilder(256);
                _ = GetWindowText(hWnd, sb, sb.Capacity);
                if (sb.ToString() == title)
                    return hWnd;
            }
            return IntPtr.Zero;
        }

        /// <summary>
        /// Close a message box with the specified title.
        /// </summary>
        /// <param name="title">The title of the message box to close.</param>
        public static void CloseMessageBox(string title)
        {
            IntPtr hWnd = GetMessageBoxHandle(title);
            if (hWnd != IntPtr.Zero)
                _ = SendMessage(hWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
        }

        /// <summary>
        /// Finds a window with the specified class name and window title.
        /// </summary>
        /// <param name="parentHandle">The handle of the parent window.</param>
        /// <param name="childAfter">The handle of the child window.</param>
        /// <param name="className">The class name of the window.</param>
        /// <param name="windowTitle">The title of the window.</param>
        /// <returns></returns>
        [LibraryImport("user32.dll", EntryPoint = "FindWindowExW", StringMarshalling = StringMarshalling.Utf16)]
        public static partial IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string className, string windowTitle);

        /// <summary>
        /// Gets the text of the specified window.
        /// </summary>
        /// <param name="hWnd">The handle of the window.</param>
        /// <param name="lpString">A <see cref="StringBuilder"/> that receives the text.</param>
        /// <param name="nMaxCount">The maximum number of characters to copy to the buffer.</param>
        /// <returns>The number of characters copied to the buffer.</returns>
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        public const UInt32 WM_CLOSE = 0x0010;

        /// <summary>
        /// Sends a message to the specified window.
        /// </summary>
        /// <param name="hWnd">The handle of the window to send the message to.</param>
        /// <param name="Msg">The message to send.</param>
        /// <param name="wParam">One of the message parameters.</param>
        /// <param name="lParam">The other message parameter.</param>
        /// <returns></returns>
        [LibraryImport("user32.dll", EntryPoint = "SendMessageW", SetLastError = true)]
        public static partial IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        /// <summary>
        /// Enumerates the child windows that belong to the specified parent window.
        /// </summary>
        /// <param name="hwnd">The handle of the parent window.</param>
        /// <param name="callback">A pointer to the callback function.</param>
        /// <param name="lParam">The application-defined value to be passed to the callback function.</param>
        /// <returns>The return value is not used.</returns>
        [LibraryImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static partial bool EnumChildWindows(IntPtr hwnd, EnumWindowsProc callback, IntPtr lParam);

        public delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

        /// <summary>
        /// Retrieves the handle of the window that contains the specified window.
        /// </summary>
        /// <param name="hWnd">The handle of the window.</param>
        /// <param name="uCmd">The command to pass to the function.</param>
        /// <returns>The handle of the window that contains the specified window.</returns>
        [LibraryImport("user32.dll", EntryPoint = "GetWindow", SetLastError = true)]
        public static partial IntPtr GetWindow(IntPtr hWnd, uint uCmd);
    }
}