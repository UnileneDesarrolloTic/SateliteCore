using System.ComponentModel.DataAnnotations;

namespace SatelliteCore.Api.Models.Request
{
     public struct DatosListarDetraccionPaginado
    {   
        public string documento { get; set; }
        [Required]
        public int Pagina { get; set; }
        [Required]
        public int RegistrosPorPagina { get; set; }

    }
}
