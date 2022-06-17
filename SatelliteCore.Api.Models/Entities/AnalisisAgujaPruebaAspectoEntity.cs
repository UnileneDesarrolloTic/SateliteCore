using System;

namespace SatelliteCore.Api.Models.Entities
{
    public struct AnalisisAgujaPruebaAspectoEntity
    {
        public string LoteAnalisis { get; set; }
        public int TipoRegistro { get; set; }
        public int Cantidad { get; set; }
        public int BaseCalculoPorcentaje { get; set; }
        public decimal? Tolerancia { get; set; }
        public int Usuario { get; set; }
        public DateTime Fecha { get; set; }

        public bool ValidarDatos()
        {
            if (string.IsNullOrEmpty(LoteAnalisis) || TipoRegistro < 0 || Cantidad < 0 || BaseCalculoPorcentaje < 1)
                return false;

            return true;
        }
    }
}
