using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Generic
{
    public struct ObtenerEstructuraFormCotizacionModel
    {
        public List<EstructuraFormularioCotizacionModel> Cabecera { get; set; }
        public List<EstructuraFormularioCotizacionModel> Detalle { get; set; }
    }
}
