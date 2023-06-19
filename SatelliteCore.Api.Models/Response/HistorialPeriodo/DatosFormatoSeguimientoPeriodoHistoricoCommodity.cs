using System;
using System.Collections.Generic;
using System.Text;
using SatelliteCore.Api.Models.Response.CompraImportacion;

namespace SatelliteCore.Api.Models.Response.HistorialPeriodo
{
    public class DatosFormatoSeguimientoPeriodoHistoricoCommodity
    {
        public IEnumerable<DatosFormatoPeriodoCommodity> Periodo { get; set; }
        public IEnumerable<DatosFormatoReporteHistoricoConsumoCommodity> PeriodoHistoricoComodity { get; set; }
    }
}
