using System;

namespace SatelliteCore.Api.Models.Entities
{
    public struct AnalisisAgujaElasticidadPerforacionEntity
    {
        public string LoteAnalisis { get; set; }
        public int TipoRegistro { get; set; }
        public decimal Uno { get; set; }
        public decimal Dos { get; set; }
        public decimal Tres { get; set; }
        public decimal Cuatro { get; set; }
        public decimal Cinco { get; set; }
        public string Estado { get; set; }
        public decimal? FuerzaPerforacion { get; set; }
        public int Usuario { get; set; }
        public DateTime Fecha { get; set; }

        public bool DatosValidos()
        {
            if (string.IsNullOrEmpty(LoteAnalisis) || TipoRegistro < 1 || Uno < 0 || Dos < 0 || Tres < 0 || Cuatro < 0 || Cinco < 0)
                return false;

            return true;
        }
    }
}
