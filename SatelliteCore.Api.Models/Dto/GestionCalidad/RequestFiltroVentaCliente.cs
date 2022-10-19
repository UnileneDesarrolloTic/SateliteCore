using System;

namespace SatelliteCore.Api.Models.Dto.GestionCalidad
{
    public struct RequestFiltroVentaCliente
    {
        public DateTime FechaInicio { get; set; }
        public DateTime FechaFin { get; set; }
        public int Cliente { get; set; }
        public string Linea { get; set; }
        public string Familia { get; set; }
        public string SubFamilia { get; set; }
        public string Item { get; set; }
        public string NumeroParte { get; set; }

        public bool ValidarDatos()
        {
            if (Cliente < 1 )
                return false;

            return true;
        }
    }
}
