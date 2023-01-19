using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Generic
{
    public struct LogTrazaEvento
    {
        public int IdEvento { get; set; }
        public string Usuario { get; set; }
        public string Opcional { get; set; }
    }
}
