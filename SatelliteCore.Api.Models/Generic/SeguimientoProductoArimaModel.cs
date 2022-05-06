using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Generic
{
    public struct SeguimientoProductoArimaModel
    {
        public List<ProductoArimaModel> Productos { get; set; }
        public List<TransitoProductoArimaModel> DetalleTransito { get; set; }
        
    }
}
