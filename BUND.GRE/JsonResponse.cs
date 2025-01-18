using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.Serialization.Json;
namespace BUND.GRE
{
    public class JsonResponse
    {
        public ResponseEnvio JsonDeserialize(string jsonString) {
            ResponseEnvio objResponseEnvio = new ResponseEnvio();
            using (var ms = new MemoryStream(Encoding.Unicode.GetBytes(jsonString)))
            {
                DataContractJsonSerializer deserializer = new DataContractJsonSerializer(typeof(ResponseEnvio));
                objResponseEnvio = (ResponseEnvio)deserializer.ReadObject(ms);
            }
            

            //ResponseEnvio objResponseEnvio = new ResponseEnvio();
            //MemoryStream ms =new MemoryStream(Encoding.Unicode.GetBytes(jsonString));
            //DataContractJsonSerializer serializer = new DataContractJsonSerializer(objResponseEnvio.GetType());
            //serializer.ReadObject(ms);
            //objResponseEnvio = (ResponseEnvio)serializer.ReadObject(ms);
            //ms.Close();
            return objResponseEnvio;
        }

    }
}
