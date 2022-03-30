using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct CompraMPArimaDetalleControlCalidad
    {
        public string Item { get; set; }
        public string NumeroOrden { get; set; }
        public decimal CantidadRecibida { get; set; }
        public string Estado { get; set; }
        public DateTime FechaPreparacion { get; set; }
        
    }
}
