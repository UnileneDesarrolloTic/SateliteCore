using System;

namespace SatelliteCore.Api.Models.Encajado
{
    public struct TransferenciaEncajadoDTO
    {
        public int Id { get; set; }
        public string UsuarioRegistro { get; set; }
        public DateTime FechaRegistro { get; set; }
        public decimal CantidadTransferida { get; set; }
        public int Avance { get; set; }
    }
}
