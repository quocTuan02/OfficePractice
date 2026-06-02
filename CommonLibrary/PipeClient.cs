using System.IO;
using System.IO.Pipes;

namespace CommonLibrary
{
    public static class PipeClient
    {
        internal const string PipeName = "MOS_Office_Pipe";

        public static void Send(PipeMessage message)
        {
            try
            {
                using (var client = new NamedPipeClientStream(".", PipeName, PipeDirection.Out))
                {
                    client.Connect(300);
                    using (var writer = new StreamWriter(client))
                    {
                        writer.Write(message.ToJson());
                        writer.Flush();
                    }
                }
            }
            catch
            {
                // MainApp may not be running — silently ignore
            }
        }
    }
}
