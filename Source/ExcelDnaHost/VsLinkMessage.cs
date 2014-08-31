
// NOTE: This file is also included (as a link) in the ExcelDnaTools project

using System;
using System.IO;
using System.Text;

namespace ExcelDna.ExcelDnaTools
{

    // Message numbers etc. are shared with the client
    // CONSIDER: Move to a shared assembly.
    //           Make nice / proper serialization, classes whatever
    enum VsLinkMessageType
    {
        EchoRequest = 1,    // Excel -> VS or VS -> Excel
        EchoReply = 2,      // Excel -> VS or VS -> Excel
        RegisterAddIn = 3,    // VS -> Excel (Reload if loaded...?) + XllPath / DllPath
    }

    // abstract or generic.... VsLinkMessage<TMessage> where TMessage: IVsLinkMessage ???
    abstract class VsLinkMessage
    {
        public virtual byte[] GetData()
        {
            return new byte[0];
        }

        public VsLinkMessageType MessageType
        {
            get
            {
                if (this is EchoRequestMessage) return VsLinkMessageType.EchoRequest;
                if (this is EchoReplyMessage)   return VsLinkMessageType.EchoReply;
                if (this is RegisterAddInMessage) return VsLinkMessageType.RegisterAddIn;
                throw new InvalidOperationException("VsLinkMessage.MessageType - unknown type");
            }
        }

        public byte[] ToBytes()
        {
            using (var ms = new MemoryStream())
            using (var sw = new BinaryWriter(ms))
            {
                sw.Write((int)MessageType);
                byte[] data = GetData();
                sw.Write(data.Length);
                sw.Write(data);
                return ms.ToArray();
            }
        }

        public static VsLinkMessage FromBytes(byte[] bytes)
        {
            if (bytes.Length < 8)
                throw new ArgumentException("VsLinkMessage is malformed - message too short");
            int messageType = BitConverter.ToInt32(bytes, 0);
            var dataLength = BitConverter.ToInt32(bytes, 4);
            if (bytes.Length - 8 != dataLength)
                throw new ArgumentException("VsLinkMessage is malformed - data length incorrect");

            switch ((VsLinkMessageType)messageType)
            {
                case VsLinkMessageType.EchoRequest:
                    return new EchoRequestMessage();
                case VsLinkMessageType.EchoReply:
                    return new EchoReplyMessage();
                case VsLinkMessageType.RegisterAddIn:
                    return new RegisterAddInMessage(bytes, 8);
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }
    }

    class EchoRequestMessage : VsLinkMessage
    {
    }

    class EchoReplyMessage : VsLinkMessage
    {
    }

    class RegisterAddInMessage : VsLinkMessage
    {
        public readonly string AddInPath;

        public RegisterAddInMessage(string addInPath)
        {
            AddInPath = addInPath;
        }

        public RegisterAddInMessage(byte[] data, int startIndex)
        {
            AddInPath = Encoding.UTF8.GetString(data, startIndex, data.Length - startIndex);
        }

        public override byte[] GetData()
        {
            return Encoding.UTF8.GetBytes(AddInPath);
        }
    }
}
