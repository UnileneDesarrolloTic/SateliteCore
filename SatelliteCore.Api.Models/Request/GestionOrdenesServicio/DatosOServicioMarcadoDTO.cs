using System;

namespace SatelliteCore.Api.Models.Request.GestionOrdenesServicio
{
    public struct DatosOServicioMarcadoDTO
    {
        public string OrdenServicio { get; set; }
        public string Transportista { get; set; }
        public DateTime? FechaRegistro { get; set; }
    }
}
