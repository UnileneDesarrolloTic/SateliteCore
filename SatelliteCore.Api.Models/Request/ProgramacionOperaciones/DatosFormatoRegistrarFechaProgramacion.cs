using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request.ProgramacionOperaciones
{
    public struct DatosFormatoRegistrarFechaProgramacion
    {   public int id { get; set; }
        public string lote { get; set; }
        public string ordenFabricacion { get; set; }
        public int cantidadProgramada { get; set; }
        public DateTime? fechaEntrega { get; set; }
        public DateTime? fechaInicio { get; set; }
        public string comentario { get; set; }
    }
}
