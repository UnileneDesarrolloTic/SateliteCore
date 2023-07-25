using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.GestioEquipoEngaste
{
    public struct DatosFormatoCabeceraInformacionEquipo
    {
        public int IdEquipo { get; set; }
        public string Descripcion { get; set; }
        public string Tipo { get; set; }
        public int IdPersona { get; set; }
        public string Estado { get; set; }
        public string NombreCompleto { get; set; }
    }
}
