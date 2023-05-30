using System;

namespace SatelliteCore.Api.Models.Request.GestionOrdenesServicio
{
    public struct DatosReporteGuiaOrdenServicioDTO
    {
        public string Documento { get; set; }
        public string Cliente { get; set; }
        public string Estado { get; set; }
        public DateTime FechaEmision { get; set; }
        public DateTime FechaSalida { get; set; }
        public DateTime FechaIngreso { get; set; }
        public DateTime FechaDespacho { get; set; }
        public DateTime FechaRetornoAlm { get; set; }
        public DateTime FechaRetornoCom { get; set; }
        public int Impresiones { get; set; }
    }
}
