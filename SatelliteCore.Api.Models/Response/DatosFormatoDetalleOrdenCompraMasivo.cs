using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatoDetalleOrdenCompraMasivo
    {
        public string NumeroOrden { get; set; }
        public string Descripcion { get; set; }
        public decimal CantidadPedida { get; set; }
        public DateTime FechaPrometida { get; set; }
        public string Item { get; set; }
    }
}
