using System.ComponentModel.DataAnnotations;

namespace SatelliteCore.Api.Models.Request
{
     public struct DatosListarDetraccionPaginado
    {
        [Required]
        public int Pagina { get; set; }

        [Required]
        public int RegistrosPorPagina { get; set; }

    }
}
