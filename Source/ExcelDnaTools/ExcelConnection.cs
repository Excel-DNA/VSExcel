using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDna.ExcelDnaTools
{
    // Counterpart of VsConnection in the ExcelDnaHost
    public class ExcelConnection
    {
        VsLinkClient _vsLinkClient;
        int _processID;

        public void AttachExcel(string pipeName)
        {
            // PipeName ends with process id
            _processID = int.Parse(pipeName.Substring(pipeName.LastIndexOf('\\') + 1));
            _vsLinkClient = new VsLinkClient(pipeName);
        }

        public Task RegisterAddIn(string pathName)
        {
            return _vsLinkClient.SendMessageAsync(new RegisterAddInMessage(pathName));
        }

        public int ProcessID { get { return _processID; } } 

        // RunMacro

    }
}
