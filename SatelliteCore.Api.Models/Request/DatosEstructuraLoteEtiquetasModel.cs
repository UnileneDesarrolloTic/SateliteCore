using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request
{
    public class DatosEstructuraLoteEtiquetasModel
    {
        public DateTime fechaProduccion { get; set; }
        public string item { get; set; }
        public string numeroParte { get; set; }
        public string marca { get; set; }
        public string descripcionLocal { get; set; }
        public string cliente { get; set; }
        public string lote { get; set; }
        public string ordenFabricacion { get; set; }
        public string transferidoflag { get; set; }
    }
}
