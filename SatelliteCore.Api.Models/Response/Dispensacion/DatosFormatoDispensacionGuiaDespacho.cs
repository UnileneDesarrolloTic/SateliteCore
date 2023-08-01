using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.Dispensacion
{
    public struct DatosFormatoDispensacionGuiaDespacho
    {
        public int Id { get; set; }
        public string EntregadoPor { get; set; }
        public string RecibidoPor { get; set; }
        public string Estado { get; set; }
        public DateTime FechaRegistro { get; set; }
        public DateTime? FechaDespacho { get; set; }
    }
}
