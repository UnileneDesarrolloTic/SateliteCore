using System.ComponentModel.DataAnnotations;

namespace SatelliteCore.Api.Models.Request
{
    public struct DatosListarCotizacionesPaginado
    {
        public string NumeroDocumento { get; set; }
        public string ClienteNombre { get; set; }
        [Required]
        public int Pagina { get; set; }
        [Required]
        public int RegistrosPorPagina { get; set; }
    }
}
