using System;
using System.IO;
using System.IO.Pipes;
using System.Threading;

namespace CommonLibrary
{
    public class PipeServer : IDisposable
    {
        public const string PipeName = PipeClient.PipeName;

        private Thread _thread;
        private volatile bool _running;

        public event Action<PipeMessage> MessageReceived;

        public void Start()
        {
            _running = true;
            _thread = new Thread(ListenLoop) { IsBackground = true, Name = "PipeServer" };
            _thread.Start();
        }

        private void ListenLoop()
        {
            while (_running)
            {
                try
                {
                    using (var server = new NamedPipeServerStream(
                        PipeName, PipeDirection.In,
                        NamedPipeServerStream.MaxAllowedServerInstances,
                        PipeTransmissionMode.Byte))
                    {
                        server.WaitForConnection();
                        using (var reader = new StreamReader(server))
                        {
                            var json = reader.ReadToEnd();
                            if (!string.IsNullOrEmpty(json))
                                MessageReceived?.Invoke(PipeMessage.FromJson(json));
                        }
                    }
                }
                catch (ThreadInterruptedException) { break; }
                catch { /* ignore connection errors */ }
            }
        }

        public void Stop()
        {
            _running = false;
            _thread?.Interrupt();
        }

        public void Dispose() => Stop();
    }
}
