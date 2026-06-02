using System;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

namespace CommonLibrary
{
    [DataContract]
    public class PipeMessage
    {
        [DataMember] public string EventType { get; set; }
        [DataMember] public string Application { get; set; }
        [DataMember] public string Detail { get; set; }
        [DataMember] public string Timestamp { get; set; }

        public PipeMessage()
        {
            Timestamp = DateTime.Now.ToString("HH:mm:ss");
        }

        public string ToJson()
        {
            var serializer = new DataContractJsonSerializer(typeof(PipeMessage));
            using (var ms = new MemoryStream())
            {
                serializer.WriteObject(ms, this);
                return Encoding.UTF8.GetString(ms.ToArray());
            }
        }

        public static PipeMessage FromJson(string json)
        {
            var serializer = new DataContractJsonSerializer(typeof(PipeMessage));
            using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(json)))
            {
                return (PipeMessage)serializer.ReadObject(ms);
            }
        }
    }
}
