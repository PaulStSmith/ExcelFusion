using System.Diagnostics;
using System.Runtime.InteropServices;

class InterceptMessages
{
    private const int WH_CBT = 5;
    private const int GW_OWNER = 5;
    private const int HCBT_CREATEWND = 3;
    private static IntPtr _hookID = IntPtr.Zero;

    [DllImport("user32.dll")]
    private static extern IntPtr SetWindowsHookEx(int idHook, CBTProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll")]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll")]
    public static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    public delegate IntPtr CBTProc(int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr GetModuleHandle(string lpModuleName);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

    //public static void Main()
    //{
    //    _hookID = SetHook(CBTCallback);
    //    // Rest of your application code

    //    // Unhook before exiting
    //    UnhookWindowsHookEx(_hookID);
    //}

    public static IntPtr SetHook(CBTProc proc)
    {
        using Process curProcess = Process.GetCurrentProcess();
        using ProcessModule curModule = curProcess.MainModule;
        return SetWindowsHookEx(WH_CBT, proc, GetModuleHandle(curModule.ModuleName), 0);
    }

    public static void Unhook()
    {
        UnhookWindowsHookEx(_hookID);
    }

    private static IntPtr CBTCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode == HCBT_CREATEWND)
        {
            // Check if the window being created is a message box from Excel
            IntPtr hwndOwner = GetWindow(wParam, GW_OWNER);
            if (IsExcelWindow(hwndOwner))
            {
                // Perform your logic here
            }
        }
        return CallNextHookEx(_hookID, nCode, wParam, lParam);
    }

    private static bool IsExcelWindow(IntPtr hWnd)
    {
        // Implement logic to check if hWnd is an Excel window
        // This might involve checking the window's class name or title
        return true; // Placeholder for actual implementation
    }
}
