using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatoInformacionCalendarioSeguimientoOC
    {
        public List<DatosFormatoCalentarioSeguimientoOC> Calendario { get; set; }
        public List<DatosFormatoDetalleCalendarioSeguimientoOC> DetalleCalendario { get; set; }
    }
}
