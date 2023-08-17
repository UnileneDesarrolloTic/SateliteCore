using SatelliteCore.Api.Models.Entities;
using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Response.AnalisisMateriaPrima
{
    public class GuardarAnalisisHebraDTO
    {
        public TBMAnalisisHebraEntity Cabecera { get; set; }
        public List<TBDAnalisisHebraEntity> Detalle { get; set; }

        public string Item { get; set; }
        public string NumeroLote { get; set; }
        public string NumeroDeParte { get; set; }
        public decimal Longitud { get; set; }
        public decimal Diametro { get; set; }
        public decimal Tension { get; set; }
    }
}
