
// NOTE: This file is also included (as a link) in the ExcelDnaTools project

using System;
using System.Collections.Generic;
using System.IO.Pipes;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelDna.ExcelDnaTools
{
    // The VsLink is used to create a VS -> Excel channel
    // (We use VS Commands and the DTE object to communicate Excel -> VS)
    // An alternative would be to use COM, and call Application.Run(...) for VS -> Excel add-in messages.
    abstract class VsLink : IDisposable
    {
        public string PipeName { get; protected set; }
        public PipeStream PipeStream { get; protected set; }
        readonly byte[] _buffer = new byte[1024];

        public bool IsConnected { get { return PipeStream.IsConnected; } }

        public Task SendMessageAsync(VsLinkMessage message)
        {
            byte[] messageBytes = message.ToBytes();
            return PipeStream.WriteAsync(messageBytes, 0, messageBytes.Length);
        }

        // TODO: Timeout and CancellationToken
        // TODO: Receive all messages / blocking receive?
        public async Task<VsLinkMessage> ReceiveMessageAsync()
        {
            var messageBytes = new List<byte>();
            do
            {
                var numBytes = await PipeStream.ReadAsync(_buffer, 0, _buffer.Length /*, cancellationToken*/);
                if (numBytes == 0)
                {
                    // End-of-stream (pipe closed?)
                    if (messageBytes.Count > 0)
                    {
                        throw new InvalidOperationException("Pipe closed while transmitting message");
                    }
                    return null;
                }
                // normal case - add to the message bytes
                messageBytes.AddRange(_buffer.Take(numBytes));
            }
            while (!PipeStream.IsMessageComplete);

            return VsLinkMessage.FromBytes(messageBytes.ToArray());
        }

        public virtual void Dispose()
        {
            PipeStream.Dispose();
        }

    }
}
