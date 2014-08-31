using System;
using System.Collections.Generic;
using System.IO.Pipes;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDna.ExcelDnaTools
{
    // Connects to the pipe id passed in
    class VsLinkClient : VsLink
    {
        public VsLinkClient(string pipeName)
        {
            var ps = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut, PipeOptions.Asynchronous);
            ps.Connect(/* timeout */);
            PipeStream = ps;
        }

    }
}
