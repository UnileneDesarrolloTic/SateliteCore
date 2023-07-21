using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.RRHH.AsignacionPersonal
{
    public struct DatosFormatoPersonasAsistencia
    {
        public int Persona { get; set; }
        public string NombreCompleto { get; set; }
        public bool Asistencia { get; set; }
        public bool Justificaciones { get; set; }
        public bool Vacaciones { get; set; }
        public bool Injustificado { get; set; }
    }
}
