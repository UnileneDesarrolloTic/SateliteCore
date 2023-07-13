using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.OCDrogueria
{
    public struct DatosFormatoMostrarOrdenCompraDrogueria
    {   
        public string ItemFinal { get; set; }
        public string Descripcion { get; set; }
        public string NumeroOrden { get; set; }
        public string Proveedor { get; set; }
        public decimal CantidadPedida { get; set; }
        public decimal CantidadRecibida { get; set; }
        public decimal CantidadPendiente { get; set; }
        public DateTime FechaPreparacion { get; set; }
        public decimal TiempoGeneral { get; set; }
        public DateTime NuevoTiempoEntrega { get; set; }
        public DateTime FechaPrometida { get; set; }
        public int DiasFalta { get; set; }
    }
}
