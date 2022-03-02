using System.ComponentModel.DataAnnotations;


namespace SatelliteCore.Api.Models.Request
{
    public struct PronosticoCompraMP
    {
        [Required]
        public string Periodo { get; set; }
        [Required]
        public string Regla { get; set; }
        [Required]
        public string Linea { get; set; }
        [Required]
        public string Familia { get; set; }
        [Required]
        public string FamiliaMP { get; set; }
        [Required]
        public string Agrupador { get; set; }
    }
}
