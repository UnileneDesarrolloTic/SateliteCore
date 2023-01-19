using System;

namespace SatelliteCore.Api.Models.Request.RRHH
{
    public struct HorasExtraExportDTO
    {
        public DateTime FechaRegistro { get; set; }
        public int IdPersona { get; set; }
        public decimal CantidadHoras { get; set; }
        public decimal H25 { get; set; }
        public decimal H35 { get; set; }
        public decimal SueldoActualLocal { get; set; }
        public decimal SueldoHora { get; set; }
        public string NombreCompleto { get; set; }
        public decimal S25 { get; set; }
        public decimal S35 { get; set; }
    }
}
