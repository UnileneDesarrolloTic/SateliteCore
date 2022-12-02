using System;

namespace SatelliteCore.Api.Models.Entities
{
    public struct TBMReclamosEntity
    {
        public string CodReclamo { get; set; }
        public int Cliente { get; set; }
        public string Estado { get; set; }
        public string UsuarioRegistro { get; set; }
        public DateTime FechaRegistro { get; set; }
    }
}
