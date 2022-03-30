using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Generic
{
    public struct SeguimientoComprasMPArima
    {

            public List<CompraMPArimaModel> Productos { get; set; }
            public List<DCompraMPArimaModel> DetalleTransito { get; set; }
            public List<CompraMPArimaDetalleControlCalidad> DetalleCalidad { get; set; }

    }
}
