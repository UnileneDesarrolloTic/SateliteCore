using System;

namespace SatelliteCore.Api.Models.Response
{
    public struct DetalleProtocoloAnalisis
    {
        public string NumeroDocumento { get; set; }
        public DateTime? FechaDocumento { get; set; }
        public DateTime? FechaVencimiento { get; set; }
        public string ClienteNombre { get; set; }
        public string ItemCodigo { get; set; }
        public string Descripcion { get; set; }
        public decimal CantidadPedida { get; set; }
        public string Lote { get; set; }
        public DateTime? FechaExpiracion { get; set; }
        public string OrdenFabricacion { get; set; }
        public string Comentarios { get; set; }
        public string TieneProtocolo { get; set; }
    }
}
