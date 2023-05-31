using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.HistorialPeriodo
{
    public struct DatosFormatoSeguimientoHistorioPeriodo
    {
        public IEnumerable<DatosFormatoPeriodo> Periodo { get; set; }
        public IEnumerable<DatosFormatoReporteHistorialPeriodoArima> PeriodoHistorico { get; set; }

    }
}
