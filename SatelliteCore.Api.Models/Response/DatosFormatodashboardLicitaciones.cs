using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatodashboardLicitaciones
    {
        public string Proceso { get; set; }
        public string DescripcionItem { get; set; }
        public string OrdenCompra { get; set; }
        public string Entrega { get; set; }
        public string Guia { get; set; }
        public string EstadoFacturacion { get; set; }
        public string Factura { get; set; }
        public decimal MontoFactura { get; set; }
        public string EstadoComercial { get; set; }
        public string EstadoLogistica { get; set; }
        public string Estado { get; set; }
        public string Usuario { get; set; }
        public decimal MontoTotalFactura { get; set; }


    }
}
