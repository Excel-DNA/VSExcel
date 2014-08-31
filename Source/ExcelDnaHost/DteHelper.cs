using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace AppToolsHost
{
    // This code from: http://www.codeproject.com/Articles/7984/Automating-a-specific-instance-of-Visual-Studio-NE
    // By Mohamed Hendawi
    class DteHelper
    {
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private const int SW_RESTORE = 9;
        [DllImport("user32.dll")]
        private static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr hWnd);

        /// <summary>
        /// Raises an instance of the Visual Studio IDE to the foreground.
        /// </summary>
        /// <param name="ide">The DTE object for the IDE you
        ///    would like to raise to the foreground</param>

        public static void ShowIDE(EnvDTE.DTE ide)
        {
            // To show an existing IDE, we get the HWND for the MainWindow
            // and do a little interop to bring the desired IDE to the
            // foreground.  I tried some of the following other potentially
            // promising approaches but could only succeed in getting the
            // IDE's taskbar button to flash (this is as designed).  Ex:
            //
            //   ide.MainWindow.Activate();
            //   ide.MainWindow.SetFocus();
            //   ide.MainWindow.Visible = true;
            //   ide.MainWindow.WindowState = EnvDTE.vsWindowState.vsWindowStateMinimize;
            //   ide.MainWindow.WindowState = EnvDTE.vsWindowState.vsWindowStateMaximize;

            System.IntPtr hWnd = (System.IntPtr)ide.MainWindow.HWnd;
            if (IsIconic(hWnd))
            {
                ShowWindowAsync(hWnd, SW_RESTORE);
            }
            SetForegroundWindow(hWnd);
            ide.MainWindow.Visible = true;
        }
    }
}
