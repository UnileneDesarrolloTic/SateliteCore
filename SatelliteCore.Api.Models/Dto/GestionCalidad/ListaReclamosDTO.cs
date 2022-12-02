using System;

namespace SatelliteCore.Api.Models.Dto.GestionCalidad
{
    public struct ListaReclamosDTO
    {
        public string CodReclamo { get; set; }
        public string NombreCliente { get; set; }
        public string TipoCliente { get; set; }
        public string Nacionalidad { get; set; }
        public string Territorio { get; set; }
        public string Estado { get; set; }
        public DateTime FechaRegistro { get; set; }
        public string UsuarioRegistro { get; set; }
        public int DiferenciaDias { get; set; }
    }
}
