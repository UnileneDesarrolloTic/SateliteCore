using System;

namespace SatelliteCore.Api.Models.Dto.GestionCalidad
{
    public struct FiltrosListaReclamosDTO
    {
        public DateTime FechaInicio { get; set; }
        public DateTime FechaFin { get; set; }
        public int Cliente { get; set; }
        public string CodReclamo { get; set; }
        public string Territorio { get; set; }
        public int Pagina { get; set; }
        public int RegistrosPorPagina { get; set; }

        public bool ValidarFiltros()
        {
            if (Pagina < 1 || RegistrosPorPagina < 1)
                return false;

            return true;
        }
    }
}
