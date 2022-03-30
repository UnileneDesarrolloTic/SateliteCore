using System;

namespace SatelliteCore.Api.Models.Entities
{
    public struct CotizacionRegistroEntity
    {
        public string Codigo { get; set; }
        public string Cotizacion { get; set; }
        public int IDFormato { get; set; }
        public int UsuarioRegistro { get; set; }
        public DateTime FechaRegistro { get; set; }
        public int UsuarioModificacion { get; set; }
        public DateTime FechaModificacion { get; set; }
    }
}
