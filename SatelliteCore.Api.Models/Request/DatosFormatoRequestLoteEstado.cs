using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request
{
    public class DatosFormatoRequestLoteEstado
    {
        public int id { get; set; }
        public string lote { get; set; }
        public string ordenFabricacion { get; set; }
        public DateTime fechaRegistro { get; set; }
        public string usuario { get; set; }
        public string estado { get; set; }
    }
}
