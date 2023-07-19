using SatelliteCore.Api.Models.Entities;
using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Response.AnalisisMateriaPrima
{
    public class GuardarAnalisisHebraDTO
    {
        public TBMAnalisisHebraEntity Cabecera { get; set; }
        public List<TBDAnalisisHebraEntity> Detalle { get; set; }
    }
}
