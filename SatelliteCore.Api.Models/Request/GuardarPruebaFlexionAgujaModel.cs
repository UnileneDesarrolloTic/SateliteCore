using System.ComponentModel.DataAnnotations;

namespace SatelliteCore.Api.Models.Request
{
    public struct GuardarPruebaFlexionAgujaModel
    {
        [Required]
        public string Lote { get; set; }

        [Required]
        public int TipoRegistro { get; set; }

        [Required]
        public int Llave { get; set; }

        [Required]
        public decimal Valor { get; set; }

        [Required]
        public int UsuarioRegistro { get; set; }
    }
}
