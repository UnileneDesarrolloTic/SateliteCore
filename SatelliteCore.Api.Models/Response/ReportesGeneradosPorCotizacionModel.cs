using System;

namespace SatelliteCore.Api.Models.Response
{
    public struct ReportesGeneradosPorCotizacionModel
    {
        public string Codigo { get; set; }
        public int IdFormato { get; set; }
        public string Formato { get; set; }
        public string UsuarioRegistro { get; set; }
        public DateTime FechaRegistro { get; set; }
        public string UsuarioModificacion { get; set; }
        public DateTime? FechaModificacion { get; set; }
    }
}
