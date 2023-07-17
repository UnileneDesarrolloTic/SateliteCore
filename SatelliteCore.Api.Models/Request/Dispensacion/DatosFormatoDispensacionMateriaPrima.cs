using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request.Dispensacion
{
    public struct DatosFormatoDispensacionMateriaPrima
    {
         public string itemTerminado { get; set; }
         public string ordenFabricacion { get; set; }
         public List<DatosFormatoDispensacionDetalleMP> detalleDispensacion { get; set; }
    }
}
