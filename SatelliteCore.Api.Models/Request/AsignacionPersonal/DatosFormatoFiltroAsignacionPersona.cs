using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request.AsignacionPersonal
{
    public struct DatosFormatoFiltroAsignacionPersona
    {
        public DateTime fechaInicio {  get; set; }
        public DateTime fechaFin { get; set; }
        public bool reporteAsistencia { get; set; }
        public bool listadoPersonal { get; set; }
    }
}
