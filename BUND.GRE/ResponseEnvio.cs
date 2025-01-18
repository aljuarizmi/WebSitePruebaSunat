using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BUND.GRE
{
    public class ResponseEnvio
    {
        public List<error> error { get; set; }
        public string codRespuesta { get; set; }
        public string arcCdr { get; set; }
        public string indCdrGenerado { get; set; }
    }
}
