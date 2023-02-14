using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.OCDrogueria
{
    public struct DatosFormatoMostrarOrdenCompraDrogueria
    {
        public string NumeroOrden { get; set; }
        public string Proveedor { get; set; }
        public decimal Cantidad { get; set; }
        public decimal CantidadRecibida { get; set; }
        public decimal CantidadPendiente { get; set; }
        public DateTime FechaPrometida { get; set; }
        public int DiferenciaFecha { get; set; }
    }
}
