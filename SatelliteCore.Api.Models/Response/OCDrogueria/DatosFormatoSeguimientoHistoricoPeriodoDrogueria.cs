using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.OCDrogueria
{
    public struct DatosFormatoSeguimientoHistoricoPeriodoDrogueria
    {
        public IEnumerable<DatosFormatoPeriodoDrogueria> PeriodoPlantilla { get; set; }
        public IEnumerable<DatosFormatoPeriodoHistoricoDrogueria> HistoricoPeriodo { get; set; }
    }
}
