using System;

namespace SatelliteCore.Api.Models.Response.TransferenciaPT
{
    public struct DatosRptTransferenciaPT
    {
        public int IdDetalle { get; set; }
        public string ControlNumero { get; set; }
        public string OrdenFabricacion { get; set; }
        public string Lote { get; set; }
        public string PedidoNumero { get; set; }
        public string Cliente { get; set; }
        public string Estado { get; set; }
        public string Item { get; set; }
        public string Descripcion { get; set; }
        public string UsuarioTraslado { get; set; }
        public DateTime? FechaTraslado { get; set; }
        public decimal CantidadTotal { get; set; }
        public decimal CantidadPendiente { get; set; }
        public decimal CantidadEnviada { get; set; }
        public string UsuarioRecepcion { get; set; }
        public DateTime? FechaRecepcion { get; set; }
        public decimal CantidadAceptada { get; set; }
        public string AlmacenCodigo { get; set; }
    }
}
