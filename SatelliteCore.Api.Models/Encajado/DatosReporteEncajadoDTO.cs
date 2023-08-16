using System;

namespace SatelliteCore.Api.Models.Encajado
{
    public struct DatosReporteEncajadoDTO
    {
        public int Id { get; set; }
        public string OrdenFabricacion { get; set; }
        public decimal CantidadTransferida { get; set; }
        public DateTime FechaRegistro { get; set; }
        public string Etapa { get; set; }
        public string UsuarioAsignado { get; set; }
        public decimal CantidadAsignada { get; set; }
        public DateTime FechaAsignada { get; set; }
        public string Estado { get; set; }
    }
}
