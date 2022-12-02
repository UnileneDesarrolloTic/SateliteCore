using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request
{
    public struct FiltrarCotizacionExportacionModel
    {   
        public string FechaInicio { get; set; }
        public string FechaFin { get; set; }
        public int Cliente { get; set; }
        public string NumeroDocumento { get; set; }
        public string Estado { get; set; }
    }
}
