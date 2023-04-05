using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.Licitaciones
{
    public struct  DatosFormatoFacturaProceso
    {
        public string NumeroDocumento { get; set; }
        public DateTime FechaDocumento { get; set; }
        public string Proceso { get; set; }
        public string OrdenCompra { get; set; }
        public decimal MontoTotal { get; set; }
        public decimal Cantidad { get; set; }
    }
}
