using System;

namespace SatelliteCore.Api.Models.Response.GestionCalidad
{
    public struct LotesFiltradosReclamo
    {
        public DateTime FechaDocumento { get; set; }
        public string CodTipoDocumento { get; set; }
        public string TipoDocumento { get; set; }
        public string NumeroDocumento { get; set; }
        public string Lote { get; set; }
        public string OrdenFabricacion { get; set; }
        public string Item { get; set; }
        public string Descripcion { get; set; }
        public decimal CantidadPedida { get; set; }
        public decimal CantidadEntregada { get; set; }
        public decimal PrecioUnitario { get; set; }
        public decimal MontoTotal { get; set; }

    }
}
