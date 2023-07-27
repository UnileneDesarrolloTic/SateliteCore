using System;

namespace SatelliteCore.Api.Models.Response.TransferenciaPT
{
    public struct PendienteRecepcionarPtDTO
    {
        public int IdDetalle { get; set; }
        public string ControlNumero { get; set; }
        public string OrdenFabricacion { get; set; }
        public string Lote { get; set; }
        public string Estado { get; set; }
        public string Item { get; set; }
        public string Descripcion { get; set; }
        public string UsuarioTraslado { get; set; }
        public DateTime FechaTraslado { get; set; }
        public decimal CantidadTotal { get; set; }
        public decimal CantidadPendiente { get; set; }
        public decimal CantidadEnviada { get; set; }
        public decimal CantidadAceptada { get; set; }
    }
}
