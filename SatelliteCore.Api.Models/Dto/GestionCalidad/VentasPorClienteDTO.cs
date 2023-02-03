using System;

namespace SatelliteCore.Api.Models.Dto.GestionCalidad
{
    public struct VentasPorClienteDTO
    {
        public string TipoDocumento { get; set; }
        public string NumeroDocumento { get; set; }
        public DateTime FechaDocumento { get; set; }
        public string Estado { get; set; }
        public string Cliente { get; set; }
        public string TipoVenta { get; set; }
        public string ComercialPedidoNumero { get; set; }
        public string Comentarios { get; set; }
        public string ItemCodigo { get; set; }
        public string NumeroDeParte { get; set; }
        public string Linea { get; set; }
        public string Familia { get; set; }
        public string SubFamilia { get; set; }
        public string Lote { get; set; }
        public string ItemSerie { get; set; }
        public string Descripcion { get; set; }
        public string UnidadCodigo { get; set; }
        public decimal CantidadEntregada { get; set; }
        public decimal PrecioUnitario { get; set; }
        public decimal Monto { get; set; }
        public decimal MontoTotal { get; set; }
    }
}
