// This code is from the CodeProject article: http://www.codeproject.com/Articles/7984/Automating-a-specific-instance-of-Visual-Studio-NE
// by Mohamed Hendawi. (Thank you!)

using System;
using System.Collections;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Threading;
using EnvDTE;
using Microsoft.Win32;

namespace MsdevManager
{
    /// <summary>
    /// Utility class to get you a list of the running instances of the Microsoft Visual 
    /// Studio IDE.  The list is obtained by looking at the system's Running Object Table (ROT)
    /// </summary>
    /// 
    /// Other ways to get a pointer to a VisualStudio instance:
    /// 
    /// EnvDTE.DTE dte = (EnvDTE.DTE) System.Runtime.InteropServices.Marshal.GetActiveObject("VisualStudio.DTE.7.1");

    public class Msdev
    {
        #region Interop imports

        [DllImport("ole32.dll")]
        public static extern int GetRunningObjectTable(int reserved, out UCOMIRunningObjectTable prot);

        [DllImport("ole32.dll")]
        public static extern int CreateBindCtx(int reserved, out UCOMIBindCtx ppbc);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private const int SW_RESTORE = 9;
        [DllImport("user32.dll")]
        private static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr hWnd);

        #endregion

        /// <summary>
        /// Get the DTE object for the instance of Visual Studio IDE that has 
        /// the specified solution open.
        /// </summary>
        /// <param name="solutionFile">The absolute filename of the solution</param>
        /// <returns>Corresponding DTE object or null if no such IDE is running</returns>
        public static EnvDTE.DTE GetIDEInstance(string solutionFile)
        {
            Hashtable runningInstances = GetIDEInstances(true);
            IDictionaryEnumerator enumerator = runningInstances.GetEnumerator();

            while (enumerator.MoveNext())
            {
                try
                {
                    _DTE ide = (_DTE)enumerator.Value;
                    if (ide != null)
                    {
                        if (ide.Solution.FullName == solutionFile)
                        {
                            return (EnvDTE.DTE)ide;
                        }
                    }
                }
                catch { }
            }

            return null;
        }

        /// <summary>
        /// Raises an instance of the Visual Studio IDE to the foreground.
        /// </summary>
        /// <param name="ide">The DTE object for the IDE you would like to raise to the foreground</param>
        public static void ShowIDE(EnvDTE.DTE ide)
        {
            // To show an existing IDE, we get the HWND for the MainWindow
            // and do a little interop to bring the desired IDE to the
            // foreground.  I tried some of the following other potentially
            // promising approaches but could only succeed in getting the
            // IDE's taskbar button to flash.  Ex:
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

        public static void ShowIDE(string solutionFile)
        {
            EnvDTE.DTE ide = Msdev.GetIDEInstance(solutionFile);
            if (ide != null)
            {
                ShowIDE(ide);
            }
            else
            {
                // To create a new instance of the IDE, opened to the selected solution we
                // could try:
                // 
                //   Type dteType = Type.GetTypeFromProgID("VisualStudio.DTE.7.1");
                //   EnvDTE.DTE dte = Activator.CreateInstance(dteType) as EnvDTE.DTE;
                //   dte.MainWindow.WindowState = EnvDTE.vsWindowState.vsWindowStateMaximize;
                //   dte.MainWindow.Visible = true;
                //   dte.Solution.Open( solutionFile.Filename );
                //
                // This works but the new devenv.exe process does not exit when you close the
                // IDE.  You could then just reattach as described and the closed IDE would 
                // quickly redisplay (possibly useful as a feature).
                //
                // Instead we lookup the path to the IDE executable in the registry and
                // just start another process.

                RegistryKey devKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\VisualStudio\\7.1\\Setup\\VS");
                string idePath = (string)devKey.GetValue("EnvironmentPath");

                System.Diagnostics.Process p = new System.Diagnostics.Process();
                p.StartInfo.RedirectStandardOutput = false;
                p.StartInfo.Arguments = solutionFile;
                p.StartInfo.FileName = idePath;
                p.StartInfo.UseShellExecute = true;
                p.Start();
            }
        }

        /// <summary>
        /// Get a table of the currently running instances of the Visual Studio .NET IDE.
        /// </summary>
        /// <param name="openSolutionsOnly">Only return instances that have opened a solution</param>
        /// <returns>A hashtable mapping the name of the IDE in the running object table to the corresponding DTE object</returns>
        public static Hashtable GetIDEInstances(bool openSolutionsOnly)
        {
            Hashtable runningIDEInstances = new Hashtable();
            Hashtable runningObjects = GetRunningObjectTable();

            IDictionaryEnumerator rotEnumerator = runningObjects.GetEnumerator();
            while (rotEnumerator.MoveNext())
            {
                string candidateName = (string)rotEnumerator.Key;
                if (!candidateName.StartsWith("!VisualStudio.DTE"))
                    continue;

                _DTE ide = rotEnumerator.Value as _DTE;
                if (ide == null)
                    continue;

                if (openSolutionsOnly)
                {
                    try
                    {
                        string solutionFile = ide.Solution.FullName;
                        if (solutionFile != String.Empty)
                        {
                            runningIDEInstances[candidateName] = ide;
                        }
                    }
                    catch { }
                }
                else
                {
                    runningIDEInstances[candidateName] = ide;
                }
            }
            return runningIDEInstances;
        }

        /// <summary>
        /// Get a snapshot of the running object table (ROT).
        /// </summary>
        /// <returns>A hashtable mapping the name of the object in the ROT to the corresponding object</returns>
        [STAThread]
        public static Hashtable GetRunningObjectTable()
        {
            Hashtable result = new Hashtable();

            int numFetched;
            UCOMIRunningObjectTable runningObjectTable;
            UCOMIEnumMoniker monikerEnumerator;
            UCOMIMoniker[] monikers = new UCOMIMoniker[1];

            GetRunningObjectTable(0, out runningObjectTable);
            runningObjectTable.EnumRunning(out monikerEnumerator);
            monikerEnumerator.Reset();

            while (monikerEnumerator.Next(1, monikers, out numFetched) == 0)
            {
                UCOMIBindCtx ctx;
                CreateBindCtx(0, out ctx);

                string runningObjectName;
                monikers[0].GetDisplayName(ctx, null, out runningObjectName);

                object runningObjectVal;
                runningObjectTable.GetObject(monikers[0], out runningObjectVal);

                result[runningObjectName] = runningObjectVal;
            }

            return result;
        }

        public static bool CompareInstances(Hashtable instances1, Hashtable instances2)
        {
            bool changed = false;
            foreach (string instances1Key in instances1.Keys)
            {
                if (!instances2.ContainsKey(instances1Key))
                {
                    changed = true;
                    break;
                }
            }

            if (!changed)
            {
                foreach (string instances2Key in instances2.Keys)
                {
                    if (!instances1.ContainsKey(instances2Key))
                    {
                        changed = true;
                        break;
                    }
                }
            }

            return changed;
        }
    }

    public class MsdevMonitorThread
    {
        public delegate void MonitorMsdevHandler();
        public event MonitorMsdevHandler Changed;

        private System.Threading.Thread m_thread = null;
        private ISynchronizeInvoke m_invokeObject = null;
        private int m_period = 2000;
        private bool m_isRunning = false;
        private bool m_openSolutionsOnly = false;

        public MsdevMonitorThread(ISynchronizeInvoke invokeObject, bool openSolutionsOnly)
        {
            m_invokeObject = invokeObject;
            m_openSolutionsOnly = openSolutionsOnly;
        }

        ~MsdevMonitorThread()
        {
            Stop();
        }

        public void Start()
        {
            m_isRunning = true;
            if (m_thread == null)
                m_thread = new System.Threading.Thread(new ThreadStart(ThreadMain));
            m_thread.Start();
        }

        public void Stop()
        {
            m_isRunning = false;
            m_thread = null;
        }

        private void ThreadMain()
        {
            // Take a snapshot of the currently running instances of Visual Studio
            // We'll also separately keep track of the solution files that each
            // instance has open at this time.  We'll use it to detect when an 
            // IDE has loaded or unloaded a solution.

            Hashtable snapshotInstances = Msdev.GetIDEInstances(m_openSolutionsOnly);
            Hashtable snapshotSolutions = new Hashtable();
            foreach (string snapshotKey in snapshotInstances.Keys)
            {
                string solutionFile = String.Empty;
                try
                {
                    EnvDTE.DTE ide = (EnvDTE.DTE)snapshotInstances[snapshotKey];
                    solutionFile = ide.Solution.FullName;
                }
                catch { }

                snapshotSolutions[snapshotKey] = solutionFile;
            }

            // We'll just keep looping in this thread, periodically checking the
            // currently running list of IDE's.  If there is any change we'll 
            // raise a Changed event.

            while (m_isRunning)
            {
                System.Threading.Thread.Sleep(m_period);
                if (Changed != null)
                {
                    Hashtable currentInstances = Msdev.GetIDEInstances(m_openSolutionsOnly);
                    bool changed = Msdev.CompareInstances(snapshotInstances, currentInstances);
                    if (changed)
                    {
                        m_invokeObject.BeginInvoke(Changed, null);
                        snapshotInstances = currentInstances;
                    }
                    else
                    {
                        foreach (string currentKey in currentInstances.Keys)
                        {
                            string prevSolutionFile = (string)snapshotSolutions[currentKey];
                            string currentSolutionFile = String.Empty;
                            try
                            {
                                EnvDTE.DTE ide = (EnvDTE.DTE)currentInstances[currentKey];
                                currentSolutionFile = ide.Solution.FullName;
                            }
                            catch { }
                            if (prevSolutionFile != currentSolutionFile)
                            {
                                m_invokeObject.BeginInvoke(Changed, null);
                                snapshotInstances = currentInstances;
                                snapshotSolutions[currentKey] = currentSolutionFile;
                                break;
                            }
                        }
                    }
                }
            }
        }
    }
}
