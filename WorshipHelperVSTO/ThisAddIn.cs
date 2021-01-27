using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools;
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

namespace WorshipHelperVSTO
{
    public partial class ThisAddIn
    {
        public static String appDataPath;// = $@"{Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)}\PowerWorship";

        // NOTE: We need a backing field to prevent the delegate being garbage collected
        private SafeNativeMethods.HookProc _keyboardProc;

        private IntPtr _hookIdKeyboard;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            _keyboardProc = KeyboardHookCallback;

            SetWindowsHooks();
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            UnhookWindowsHooks();
        }

        private void SetWindowsHooks()
        {
            uint threadId = (uint)SafeNativeMethods.GetCurrentThreadId();

            _hookIdKeyboard =
                SafeNativeMethods.SetWindowsHookEx(
                    (int)SafeNativeMethods.HookType.WH_KEYBOARD,
                    _keyboardProc,
                    IntPtr.Zero,
                    threadId);
        }

        private void UnhookWindowsHooks()
        {
            SafeNativeMethods.UnhookWindowsHookEx(_hookIdKeyboard);
        }

        private IntPtr KeyboardHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            Application app = Globals.ThisAddIn.Application;

            if (nCode >= 0)
            {
                // Various attempts to decode this value as a proper struct were unsuccessful, with memory violation errors...
                bool keyUp = (lParam.ToInt32() & 0xC0000000) != 0 && nCode == 0;

                bool presenting = false;
                if (app.Presentations.Count > 0) // if there is a presentation, there is an active presentation
                {
                    foreach (DocumentWindow win in app.ActivePresentation.Windows)
                    {
                        // There is probably a better way...
                        if (win.Caption.Contains("Presenter View") && win.Active == Microsoft.Office.Core.MsoTriState.msoTrue)
                        {
                            presenting = true;
                        }
                    }
                }

                if (keyUp && presenting)
                {
                    char keyPressed = (char)wParam.ToInt32();
                    if (keyPressed == 'A')
                    {
                        // Insert song or scripture
                        new AddContentLiveForm().Show();
                    } else if (keyPressed == 'L')
                    {
                        // Go to logo
                        System.Windows.Forms.MessageBox.Show($"Key detected: Go to Logo");
                    }
                    
                }
                
            }

            return SafeNativeMethods.CallNextHookEx(
                _hookIdKeyboard,
                nCode,
                wParam,
                lParam);
        }

        [StructLayout(LayoutKind.Sequential)]
        private class KBDLLHOOKSTRUCT
        {
            public int vkCode;
            public int scanCode;
            public int flags;
            public int time;
            public int dwExtraInfo;
        }

        public static int FLAG_RELEASED = 0x80;

        internal static class SafeNativeMethods
    {
        public delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);

            

        public enum HookType
        {
            WH_KEYBOARD = 2
        }

        public enum WindowMessages : uint
        {
            WM_KEYDOWN = 0x0100,
            WM_KEYFIRST = 0x0100,
            WM_KEYLAST = 0x0108,
            WM_KEYUP = 0x0101,
            WM_SYSDEADCHAR = 0x0107,
            WM_SYSKEYDOWN = 0x0104,
            WM_SYSKEYUP = 0x0105
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr SetWindowsHookEx(
            int idHook,
            HookProc lpfn,
            IntPtr hMod,
            uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr CallNextHookEx(
            IntPtr hhk,
            int nCode,
            IntPtr wParam,
            IntPtr lParam);

        [DllImport("kernel32", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int GetCurrentThreadId();
    }

    // Called from ribbon constructor
    public static void PreInitialize()
        {
            // Get the assembly information
            System.Reflection.Assembly assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly();

            // CodeBase is the location of the DLL
            Uri uriCodeBase = new Uri(assemblyInfo.CodeBase);
            appDataPath = Path.GetDirectoryName(uriCodeBase.LocalPath.ToString()) + "\\Data";
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
