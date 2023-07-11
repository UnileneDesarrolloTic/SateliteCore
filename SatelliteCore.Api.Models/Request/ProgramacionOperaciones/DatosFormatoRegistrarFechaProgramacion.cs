using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request.ProgramacionOperaciones
{
    public struct DatosFormatoRegistrarFechaProgramacion
    {   
        public string ordenFabricacion { get; set; }
        public DateTime fecha { get; set; }
        public string tipoFecha { get; set; }
        public string comentario { get; set; }
    }
}
