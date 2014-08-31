using System;
namespace ExcelDna.ExcelDnaTools
{
    public class VsIde
    {
        EnvDTE.DTE _dte;
        readonly string _pipeName;

        public VsIde(string pipeName)
        {
            _pipeName = pipeName;
        }

        public void Show()
        {
            if (_dte == null)
            {
                // Check whether there are any VS processes running.
                // TODO: Show dialog with instances to connect to. For now make a new one...?
                // var instances = MsdevManager.Msdev.GetIDEInstances(false);
                // For now, just create a new instance
                Type type = Type.GetTypeFromProgID("VisualStudio.DTE");
                object dte = Activator.CreateInstance(type, true);
                _dte = (EnvDTE.DTE)dte;

                //System.IServiceProvider serviceProvider = new Microsoft.VisualStudio.Shell.ServiceProvider(_dte as Microsoft.VisualStudio.OLE.Interop.IServiceProvider);
                //IVsShell shell = serviceProvider.GetService(typeof(SVsShell)) as IVsShell;
                //if (shell == null)
                //{
                //    return;
                //}

                //IEnumPackages enumPackages;
                //int result = shell.GetPackageEnum(out enumPackages);

                //IVsPackage package = null;
                //Guid PackageToBeLoadedGuid =
                //    new Guid("10e82a35-4493-43be-b6d3-228399509924");
                //shell.LoadPackage(ref PackageToBeLoadedGuid, out package);

                _dte.MainWindow.WindowState = EnvDTE.vsWindowState.vsWindowStateNormal;

                // Load the AppToolsClient into Visual Studio, and connect the pipes.
            }
            _dte.MainWindow.Visible = true;
            // Now bring the _dte to the front (this might be done by the AppToolsClient?)
            if (_dte.MainWindow.WindowState == EnvDTE.vsWindowState.vsWindowStateMinimize)
            {
                _dte.MainWindow.WindowState = EnvDTE.vsWindowState.vsWindowStateNormal;
            }
            _dte.MainWindow.Activate();
            _dte.MainWindow.SetFocus(); 
            _dte.ExecuteCommand("ExcelDna.AttachExcel", _pipeName);
        }

        public void Hide()
        {
            _dte.MainWindow.Visible = false;
        }

        public void Shutdown()
        {
            _dte.Quit();
        }
    }
}
