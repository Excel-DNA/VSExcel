using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Debugger.Interop;
using Microsoft.VisualStudio.Shell.Interop;
using EnvDTE;
using EnvDTE80;
using System.Diagnostics;
using Microsoft.VisualStudio.Shell;


namespace ExcelDna.ExcelDnaTools
{
    // Projects that do this:
    // * https://github.com/erlandranvinge/ReAttach
    // * https://github.com/ashmind/AttachToAnything

    // Alternative with more control is in IVsSolutionBuildManager (package can implement this interface, and implement DebugLaunch):
    // http://msdn.microsoft.com/en-us/library/microsoft.visualstudio.shell.interop.ivssolutionbuildmanager.debuglaunch.aspx

    // And the whole sorry list of ways to attach:
    // http://stackoverflow.com/questions/15524038/visual-studio-how-to-get-idebugengine2-from-vs-package-except-ivsloader

    class DebugManager : IVsDebuggerEvents, IDebugEventCallback2
    {
        readonly ExcelDnaToolsPackage _package;
        readonly ExcelConnection _connection;
        readonly IVsDebugger _debugger; // IVsDebugger2 ???
        readonly DTE _dte;
        readonly Debugger2 _dteDebugger;
        readonly uint _debuggerEventsCookie;

        public DebugManager(ExcelDnaToolsPackage package, ExcelConnection connection)
        {
            _package = package;
            _connection = connection;
            var packageServiceProvider = (IServiceProvider)package;
             _debugger = packageServiceProvider.GetService(typeof(SVsShellDebugger)) as IVsDebugger;
            // var dgr = Package.GetGlobalService(typeof(SVsShellDebugger)) ;
            // _debugger = dgr as IVsDebugger;
            _dte = packageServiceProvider.GetService(typeof(SDTE)) as DTE;
            if (_dte != null)
            {
                _dteDebugger = _dte.Debugger as Debugger2;
            }

            if (_package == null || _debugger == null || _dte == null || _dteDebugger == null)
            {
                Debug.Fail("DebugManager setup failed");
                return;
            }

            if (_debugger.AdviseDebuggerEvents(this, out _debuggerEventsCookie) != VSConstants.S_OK)
            {
                Debug.Fail("DebugManager setup failed");
            }

            if (_debugger.AdviseDebugEventCallback(this) != VSConstants.S_OK)
            {
                Debug.Fail("DebugManager setup failed");
            }
        }

        public void AttachDebuggerToExcel()
        {
            var process = _dteDebugger.LocalProcesses.OfType<Process2>().Where(p => p.ProcessID == _connection.ProcessID).FirstOrDefault();
            if (process == null)
            {
                Debug.Fail("Could not find Excel process!?");
            }

            try
            {
                process.Attach2("Managed (v4.5, v4.0)");  // Maybe "Managed (v4.0)" under VS 2010??? or VsConstants.DebugEngineGuids.Managedv4...
            }
            catch (Exception e)
            {
                Debug.Fail("Couldn't launch debugger: " + e);
            }
            // TODO: We might need to set DBGLAUNCH_DetachOnStop or something
            //          IVsDebugLaunch debugger = GetService(typeof(SVsDebugLaunch)) as IVsDebugLaunch;
            //          debugger.LaunchDebugTargets(
            //          IVsDebugLaunch.DebugLaunch
            //          Use DBGLAUNCH_DetachOnStop
            // Also: IVsProjectDebugTargetProvider ???
        }


        public int OnModeChange(DBGMODE dbgmodeNew)
        {
            return VSConstants.S_OK;
        }

        public int Event(IDebugEngine2 pEngine, IDebugProcess2 pProcess, IDebugProgram2 pProgram, IDebugThread2 pThread, IDebugEvent2 pEvent, ref Guid riidEvent, uint dwAttrib)
        {
            return VSConstants.S_OK;
        }
    }
}
