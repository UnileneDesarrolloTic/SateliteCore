using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.RRHH.AsignacionPersonal
{
    public struct DatosFormatoPersonasAsistencia
    {
        public int Persona { get; set; }
        public string NombreCompleto { get; set; }
        public int Lectora { get; set; }
    }
}
