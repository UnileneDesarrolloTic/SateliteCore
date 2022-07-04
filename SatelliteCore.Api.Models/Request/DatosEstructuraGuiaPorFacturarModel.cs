using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request
{
    public class DatosEstructuraGuiaPorFacturarModel
    {
        public int destinatario { get; set; }
        public string Territorio { get; set; }
        public string FechaInicio { get; set; }
        public string FechaFin { get; set; }
        public string Tipo { get; set; }

    }
}
