using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request.Dispensacion
{
    public struct DatosFormatoFiltroOrdenFabricacion
    {
        public DateTime fechaInicio { get; set; }
        public DateTime fechaFinal { get; set; }
        public string lote { get; set; }
        public string ordenFabricacion { get; set; }
        public string estado { get; set; }

    }
}
