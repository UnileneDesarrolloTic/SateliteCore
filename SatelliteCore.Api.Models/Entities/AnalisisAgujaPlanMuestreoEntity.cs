using System;

namespace SatelliteCore.Api.Models.Entities
{
    public struct AnalisisAgujaPlanMuestreoEntity
    {
        public string LoteAnalisis { get; set; }
        public int Cantidad { get; set; }
        public int UndMuestrear { get; set; }
        public int UndMuestrearI { get; set; }
        public int UndMuestrearIII { get; set; }
        public int CajasMuestrear { get; set; }
        public string StatusFlexion { get; set; }
        public int Usuario { get; set; }
        public DateTime Fecha { get; set; }

        public bool ValidarDatos()
        {
            if (string.IsNullOrEmpty(LoteAnalisis) || string.IsNullOrEmpty(StatusFlexion) || Cantidad < 1 || UndMuestrear < 1 || UndMuestrearI < 1 || UndMuestrearIII < 1 || CajasMuestrear < 1)
                return false;

            return true;
        }
    }
}
