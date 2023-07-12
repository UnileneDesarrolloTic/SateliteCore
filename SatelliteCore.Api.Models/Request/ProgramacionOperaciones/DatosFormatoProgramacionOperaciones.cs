using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request.ProgramacionOperaciones
{
    public struct DatosFormatoProgramacionOperaciones
    {   
            public string gerencia { get; set; }
            public List<int> agrupador { get; set; }
            public DateTime? fechaInicio { get; set; }
            public DateTime? fechaFinal { get; set; }
            public string lote { get; set; }
            public string ordenFabricacion { get; set; }
            public string venta { get; set; }
            public string estado { get; set; }
    }
}
