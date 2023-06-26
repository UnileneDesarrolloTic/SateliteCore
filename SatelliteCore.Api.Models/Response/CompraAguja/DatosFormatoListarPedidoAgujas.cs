using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.CompraAguja
{
    public struct DatosFormatoListarPedidoAgujas
    {
        public string Itemcomponente { get; set; }
        public string ItemTerminado { get; set; }
        public decimal Cantidadpedida { get; set; }
        public string Pedido { get; set; }
        public string Cliente { get; set; }
        public decimal ComponenteCantidadNeta { get; set; }
        public decimal Mermapura { get; set; }
        public decimal TotalRequerida { get; set;}
        public DateTime FechaPreparacion { get; set; }
        public string Estado { get; set; }
    }
}
