using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request.ProgramacionOperaciones
{
    public struct DatosFormatoRegistrarFechaProgramacion
    {   
        public DateTime? programacionInicio { get; set; }
        public DateTime? programacionEntrega { get; set; }
        public string ordenFabricacion { get; set; }
        public DateTime? fechaEntrega { get; set; }
        public DateTime? fechaInicio { get; set; }
        public string tipoFechaEntrega { get; set; }
        public string tipoFechaInicio { get; set; }
        public string comentarioEntrega { get; set; }
        public string comentarioInicio { get; set; }
    }
}
