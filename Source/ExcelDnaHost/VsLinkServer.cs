using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.IO.Pipes;
using ExcelDna.Integration;
using System.Threading;

namespace ExcelDna.ExcelDnaTools
{
    // More Pipe info: http://www.codeproject.com/Tips/492231/Csharp-Async-Named-Pipes

    // Provides the server-side named pipe server when linked to a Vs instance
    // TODO: Shutdown?
    class VsLinkServer : VsLink, IDisposable
    {
        readonly Action<VsLinkMessage> _messageHandler;
        Thread _listenerThread;

        // messageHandler will be called from the listener thread, so should not block
        public VsLinkServer(Action<VsLinkMessage> messageHandler)
        {
            PipeName = @"ExcelDna\VsLink\" + Process.GetCurrentProcess().Id;
            // We create an anonymous named pipe, and then return the pipeName to pass to Vs.
            var ps = new NamedPipeServerStream(PipeName, PipeDirection.InOut, 1, PipeTransmissionMode.Message, PipeOptions.Asynchronous);
            _messageHandler = messageHandler;
            PipeStream = ps;
            _listenerThread = new Thread(Run);
            _listenerThread.Start();
        }

        public void Run(object _unused_)
        {
            ((NamedPipeServerStream)PipeStream).WaitForConnection();
            while (true)
            {
                var message = ReceiveMessageAsync().Result;
                if (message == null)
                {
                    return;
                }
                _messageHandler(message);
            }
        }

        public override void Dispose()
        {
            if (_listenerThread != null)
            {
                // TODO: Does this shut down the blocking WaitForConnection and Read calls cleanly?
                _listenerThread.Abort();
                _listenerThread = null;
            }
            base.Dispose();
        }
    }
}
