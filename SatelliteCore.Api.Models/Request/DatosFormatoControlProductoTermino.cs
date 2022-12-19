using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request
{
    public class DatosFormatoControlProductoTermino
    {
        public string fechaanalisis { get; set; }
        public string Numerolote { get; set; }
        public List<DatosFormatosTablaAControlProcesos> TablaLongitud { get; set; }
        public List<DatosFormatosTablaBControlProcesos> TablaResistencia { get; set; }

    }
}
