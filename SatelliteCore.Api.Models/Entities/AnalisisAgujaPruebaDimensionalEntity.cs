using System;

namespace SatelliteCore.Api.Models.Entities
{
    public struct AnalisisAgujaPruebaDimensionalEntity
    {
        public string LoteAnalisis { get; set; }
        public int TipoRegistro { get; set; }
        public int Cantidad { get; set; }
        public int BaseCalculoEstado { get; set; }
        public decimal Tolerancia { get; set; }
        public string DescripcionAux { get; set; }
        public int? CantidadAux { get; set; }
        public int Usuario { get; set; }
        public DateTime Fecha { get; set; }

        public bool ValidarDatos()
        {
            if (string.IsNullOrEmpty(LoteAnalisis) || TipoRegistro < 1 || Cantidad < 0 || Tolerancia < 0)
                return false;

            return true;
        }
    }
}
