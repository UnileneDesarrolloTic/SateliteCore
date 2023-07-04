using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Request
{
    public struct DatosReporteProtocoloAnalisis
    {   
        public int idioma { get; set; }
        public List<string> OrdenesFabricacion { get; set; }
    }
}
