using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatoInformacionResultadoProtocolo
    { 
        public string TABLA { get; set; }
        public int SECUENCIA { get; set; }
        public decimal COL_1 { get; set; }
        public decimal COL_2 { get; set; }

    }
}
