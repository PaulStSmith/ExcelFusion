using System.Diagnostics;
using System.Runtime.InteropServices;

// Suppressing warning about unnecessary suppression to keep the other suppression active.
#pragma warning disable IDE0079

// Disabling: Use 'LibraryImportAttribute' instead of 'DllImportAttribute' to generate P/Invoke marshalling code at compile time
// Reason: Honestly, I could never make LibraryImport work in any of my projects, so until I can figure it out, I'm disabling this warning.
#pragma warning disable SYSLIB1054

// Specify marshaling for P/Invoke string arguments
// Reason: Ensures correct character set handling for Windows API calls.
#pragma warning disable CA2101 

/// <summary>
/// Provides methods to intercept and handle Windows messages, specifically for CBT (Computer-Based Training) hooks.
/// </summary>
class InterceptMessages
#pragma warning restore IDE0079 // Remove unnecessary suppression
{
    /// <summary>
    /// Specifies the hook type for CBT (Computer-Based Training) hooks.
    /// </summary>
    private const int WH_CBT = 5;

    /// <summary>
    /// Stores the hook handle. Currently initialized to <see cref="IntPtr.Zero"/>.
    /// </summary>
    private static readonly IntPtr _hookID = IntPtr.Zero;

    /// <summary>
    /// Installs an application-defined hook procedure into a hook chain.
    /// </summary>
    /// <param name="idHook">Type of hook procedure to be installed.</param>
    /// <param name="lpfn">Pointer to the hook procedure.</param>
    /// <param name="hMod">Handle to the DLL containing the hook procedure.</param>
    /// <param name="dwThreadId">Identifier of the thread with which the hook procedure is to be associated.</param>
    /// <returns>If successful, returns the handle to the hook procedure; otherwise, <see cref="IntPtr.Zero"/>.</returns>
    [DllImport("user32.dll")]
    private static extern IntPtr SetWindowsHookEx(int idHook, CBTProc lpfn, IntPtr hMod, uint dwThreadId);

    /// <summary>
    /// Removes a hook procedure installed in a hook chain by <see cref="SetWindowsHookEx"/>.
    /// </summary>
    /// <param name="hhk">Handle to the hook to be removed.</param>
    /// <returns><c>true</c> if successful; otherwise, <c>false</c>.</returns>
    [DllImport("user32.dll")]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    /// <summary>
    /// Passes the hook information to the next hook procedure in the current hook chain.
    /// </summary>
    /// <param name="hhk">Handle to the current hook.</param>
    /// <param name="nCode">Hook code passed to the current hook procedure.</param>
    /// <param name="wParam">wParam value passed to the current hook procedure.</param>
    /// <param name="lParam">lParam value passed to the current hook procedure.</param>
    /// <returns>The value returned by the next hook procedure in the chain.</returns>
    [DllImport("user32.dll")]
    public static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    /// <summary>
    /// Retrieves a handle to the top-level window whose class name and window name match the specified strings.
    /// </summary>
    /// <param name="lpClassName">Class name of the window to find.</param>
    /// <param name="lpWindowName">Window name (title) of the window to find.</param>
    /// <returns>If successful, returns a handle to the window; otherwise, <see cref="IntPtr.Zero"/>.</returns>
    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    /// <summary>
    /// Represents the callback method for CBT hook procedures.
    /// </summary>
    /// <param name="nCode">Hook code.</param>
    /// <param name="wParam">Additional information.</param>
    /// <param name="lParam">Additional information.</param>
    /// <returns>An <see cref="IntPtr"/> result, depending on the hook processing.</returns>
    public delegate IntPtr CBTProc(int nCode, IntPtr wParam, IntPtr lParam);

    /// <summary>
    /// Retrieves a module handle for the specified module.
    /// </summary>
    /// <param name="lpModuleName">Name of the loaded module.</param>
    /// <returns>If successful, returns a handle to the specified module; otherwise, <see cref="IntPtr.Zero"/>.</returns>
    [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr GetModuleHandle(string lpModuleName);

    /// <summary>
    /// Retrieves a handle to a window that has the specified relationship (owner or child) to the specified window.
    /// </summary>
    /// <param name="hWnd">Handle to a window.</param>
    /// <param name="uCmd">Specifies the relationship to retrieve.</param>
    /// <returns>If successful, returns a handle to the related window; otherwise, <see cref="IntPtr.Zero"/>.</returns>
    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

    /// <summary>
    /// Sets a CBT hook with the specified callback procedure.
    /// </summary>
    /// <param name="proc">The callback procedure to associate with the hook.</param>
    /// <returns>The handle to the installed hook procedure.</returns>
    public static IntPtr SetHook(CBTProc proc)
    {
        using var curProcess = Process.GetCurrentProcess();
        using var curModule = curProcess.MainModule!;
        return SetWindowsHookEx(WH_CBT, proc, GetModuleHandle(curModule.ModuleName), 0);
    }

    /// <summary>
    /// Removes the previously installed CBT hook.
    /// </summary>
    public static void Unhook()
    {
        UnhookWindowsHookEx(_hookID);
    }
}
