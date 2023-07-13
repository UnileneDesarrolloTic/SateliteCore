using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request.ProgramacionOperaciones
{
    public struct  DatosFormatoDividirRegistroProgramacion
    {
        public string ordenFabricacion { get; set; }
        public string lote { get; set; }
        public int cantidadProgramada { get; set; }
        public List<DatosFormatoDivisionProgramacion> divisionProgramacion { get; set; }
    }
}
