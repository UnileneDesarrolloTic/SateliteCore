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
        public string FacturaNumero { get; set; }
        public DateTime? FacturaFecha { get; set; }
        public string EstadoFacturacion { get; set; }
        public DateTime FechaDocumento { get; set; }
        public string NumeroOrden { get; set; }
        public string Descripcion { get; set; }
        public decimal Cantidad { get; set; }
        public string LicitacionNumeroProceso { get; set; }
        public string ReprogramacionPuntoPartida { get; set; }
        public string RetornoAlmacen { get; set; }
        public string RetornoComercial { get; set; }
        public DateTime FechaRecepcion { get; set; }
        public DateTime? FechaRetornoComercial { get; set; }
        public DateTime? FechaRetornoAlmacen { get; set; }
        public string DepartamentoCliente { get; set; }


    }
}
