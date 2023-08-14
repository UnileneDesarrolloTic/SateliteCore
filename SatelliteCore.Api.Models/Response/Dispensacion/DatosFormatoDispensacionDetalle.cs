using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.Dispensacion
{
    public struct DatosFormatoDispensacionDetalle
    {
        public List<DatosFormatoDispensacionRecetaDetalle> DetalleDispensacion { get; set; }
        public List<SubFamiliaDispensacion> SubFamilia { get; set; }
    }
}
