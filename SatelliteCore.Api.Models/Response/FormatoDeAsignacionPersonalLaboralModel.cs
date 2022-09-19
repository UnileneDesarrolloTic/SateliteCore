using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct FormatoDeAsignacionPersonalLaboralModel
    {
        public int idEmpleado { get; set; }
        public string TipoPlanilla { get; set;  }
        public string NombreCompleto { get; set; }
        public string PROGRAMACION { get; set; }
        public int TipoHorario { get; set; }
        public string AREA_HORARIO { get; set; }
        public string HORA { get; set; }
        public bool isSelected { get; set; }

    }
}
