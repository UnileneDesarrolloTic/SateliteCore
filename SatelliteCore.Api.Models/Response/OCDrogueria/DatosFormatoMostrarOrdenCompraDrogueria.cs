using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.OCDrogueria
{
    public struct DatosFormatoMostrarOrdenCompraDrogueria
    {   
        public string NumeroOrden { get; set; }
        public string Proveedor { get; set; }
        public decimal CantidadPedida { get; set; }
        public decimal CantidadRecibida { get; set; }
        public decimal CantidadKardex { get; set; }
        public decimal CantidadTransito { get; set; }
        public DateTime FechaPreparacion { get; set; }
        public int TiempoGeneral { get; set; }
        public DateTime NuevoTiempoEntrega { get; set; }
    }
}
