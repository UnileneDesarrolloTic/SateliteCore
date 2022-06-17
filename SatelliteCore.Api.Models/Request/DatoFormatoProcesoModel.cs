using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request
{
   public class DatoFormatoProcesoModel
    {
        public int Cliente { get; set; }
        public string Proceso { get; set; }
        public List<DatosFormatoDetalleProceso> DetalleProceso { get; set; }
    }
}
