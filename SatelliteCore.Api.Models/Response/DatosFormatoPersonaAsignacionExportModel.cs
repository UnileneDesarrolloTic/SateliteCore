using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatoPersonaAsignacionExportModel
    {   
        public int Persona { get; set; }
        public int IdAsignacion { get; set;  }
        public string NombreCompleto { get; set; }
        public string NombreArea { get; set; }
        public DateTime FechaAsignacion { get; set; }
        public DateTime FechaReAsignacion { get; set; }
        public string Estado { get; set; }
    }
}
