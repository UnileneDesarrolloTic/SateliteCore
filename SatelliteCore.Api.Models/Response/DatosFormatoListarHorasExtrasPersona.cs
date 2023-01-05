using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatoListarHorasExtrasPersona
    {
        public int IdCabecera { get; set; }
        public int IdArea { get; set; }
        public string NombreArea { get; set; }
        public DateTime FechaRegistro { get; set; }
        public string TipoPersona { get; set; }
        public string Justificacion { get; set; }
        public string Estado { get; set; }
        public string UsuarioAprobacion { get; set; }
        public DateTime FechaAprobacion { get; set; }
    }
}
