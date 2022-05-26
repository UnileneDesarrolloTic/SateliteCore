using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Response
{
    public struct FormatoReporteGuiaRemisionesModel
    {

        public List<CReporteGuiaRemisionModel> CabeceraReporteGuiaRemision { get; set; }
        public List<DReportGuiaRemisionModel> DetalleReporteGuiaRemision { get; set; }
    }
}
