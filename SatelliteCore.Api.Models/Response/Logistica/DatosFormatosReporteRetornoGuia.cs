using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.Logistica
{
    public struct DatosFormatosReporteRetornoGuia
    {
        public string SerieNumero { get; set; }
        public string GuiaNumero { get; set; }
        public string Cliente { get; set; }
        public DateTime FechaDocumento { get; set; }
        public string NumeroOrden { get; set; }
        public string LicitacionNumeroProceso { get; set; }
        public string ReprogramacionPuntoPartida { get; set; }
        public string RetornoAlmacen { get; set; }
        public string RetornoComercial { get; set; }
        public DateTime? FechaRecepcion { get; set; }
        public DateTime? FechaRetornoComercial { get; set; }
        public DateTime? FechaRetornoAlmacen { get; set; }
        public int DiasAtrasoAlmacen { get; set; }
        public int DiasAtrasoComercial { get; set; }
        public string Destino { get; set; }
        public string Transportista { get; set; }
        public string Provincia { get; set; }
        public decimal Monto { get; set; }
        public decimal MontoOrdenCompra { get; set; }


    }
}
