using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Dto.GestionCalidad
{
    public struct RequestLotesDetalleDTO
    {
        public List<string> Lotes { get; set; }
        public List<string> OrdenesFabricacion { get; set; }
    }
}
