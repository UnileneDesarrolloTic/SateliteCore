using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.CompraAguja
{
    public class DatosInformacionGeneralReporteCompraArimaAgujas
    {
        public IEnumerable<DatosFormatoListadoSeguimientoCompraAguja> DetalleInformacionAguja { get; set; }
        public IEnumerable<DatosFormatoCantidadTotalAgujas> Total { get; set; }
    }
}
