using System;

namespace SatelliteCore.Api.Models.Dto.GestionCalidad
{
    public struct RequestFiltroVentaCliente
    {
        public DateTime? FechaInicio { get; set; }
        public DateTime? FechaFin { get; set; }
        public int Cliente { get; set; }
        public string Linea { get; set; }
        public string Familia { get; set; }
        public string SubFamilia { get; set; }
        public string Item { get; set; }
        public string NumeroParte { get; set; }
        public string Lote { get; set; }

        public bool ValidarDatos()
        {
            if (Cliente < 1 && string.IsNullOrEmpty(Lote) )
                return false;

            if (string.IsNullOrEmpty(Lote) && Cliente != 0 && (FechaInicio is null || FechaFin is null))
                return false;

            return true;
        }
    }
}
