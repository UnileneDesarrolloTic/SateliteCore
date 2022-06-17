using SatelliteCore.Api.Models.Entities;
using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Dto.AnalisisAgujas
{
    public struct PruebaAspectoYObservacionesDTO
    {
        public List<AnalisisAgujaPruebaAspectoEntity> Pruebas { get; set; }
        public string Observaciones { get; set; }        
    }
}
