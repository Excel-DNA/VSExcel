using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;

namespace ExcelDna.ExcelDnaTools
{
    static class VsConnection
    {
        static readonly VsLinkServer _vsLinkServer;
        static readonly VsIde _vsIde;

        static VsConnection()
        {
            _vsLinkServer = new VsLinkServer(VsLinkMessageHandler);
            _vsIde = new VsIde(_vsLinkServer.PipeName);
        }

        // Called on the VsLinkServer thread - should not block
        public static void VsLinkMessageHandler(VsLinkMessage message)
        {
            var registerXllMessage = message as RegisterAddInMessage;
            if (registerXllMessage != null)
            {
                if (registerXllMessage.AddInPath.EndsWith(".xll"))
                {
                    ExcelAsyncUtil.QueueAsMacro(() => ExcelIntegration.RegisterXLL(registerXllMessage.AddInPath));
                }
                else if (registerXllMessage.AddInPath.EndsWith(".dll"))
                {
                    ExcelAsyncUtil.QueueAsMacro(() => AddInLoader.RegisterDll(registerXllMessage.AddInPath));
                }
            }
        }

        public static void Show()
        {
            var Application = ExcelDnaUtil.Application as Application;
            var oldStatus = Application.StatusBar;
            var oldCursor = Application.Cursor;
            try
            {
                Application.StatusBar = "Loading Visual Studio ...";
                Application.Cursor = XlMousePointer.xlWait;
                _vsIde.Show();
            }
            catch (Exception ex)
            {
            }
            finally
            {
                Application.StatusBar = oldStatus;
                Application.Cursor = oldCursor;
            }
        }

        public static void Hide()
        {
            _vsIde.Hide();
        }

        public static void Shutdown()
        {
            _vsIde.Shutdown();
            _vsLinkServer.Dispose();
        }

    }
}
