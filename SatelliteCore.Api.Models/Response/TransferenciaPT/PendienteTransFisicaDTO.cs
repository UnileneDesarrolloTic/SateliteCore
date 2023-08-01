using System;

namespace SatelliteCore.Api.Models.Response.TransferenciaPT
{
    public struct PendienteTransFisicaDTO
    {
        public int Id { get; set; }
        public string ControlNumero { get; set; }
        public string OrdenFabricacion { get; set; }
        public string Lote { get; set; }
        public string Estado { get; set; }
        public DateTime FechaPreparacion { get; set; }
        public DateTime FechaVencimiento { get; set; }
        public decimal CantidadRecibida { get; set; }
        public string Item { get; set; }
        public string ItemDescripcion { get; set; }
        public string PedidoNumero { get; set; }
        public string Cliente { get; set; }
        public string AlmacenCodigo { get; set; }
        public decimal? CantidadEntregada { get; set; }
        public decimal? CantidadAceptada { get; set; }
        public decimal? CantidadPendiente { get; set; }
    }
}
