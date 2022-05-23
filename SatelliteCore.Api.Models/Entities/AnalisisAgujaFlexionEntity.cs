using System;

namespace SatelliteCore.Api.Models.Entities
{
    public struct AnalisisAgujaFlexionEntity
    {
        public int IdAnalisis { get; set; }
        public string Lote { get; set; }
        public int TipoRegistro { get; set; }
        public int Llave { get; set; }
        public decimal Valor { get; set; }
        public int UsuarioRegistro { get; set; }
        public DateTime FechaRegistro { get; set; }
    }
}
