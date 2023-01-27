using System.ComponentModel.DataAnnotations;

namespace SatelliteCore.Api.Models.Request
{
    public class DatosFormatoDetalleFormatoProtocolo
    {
        public int orden { get; set; }
        public string descripcionLocal { get; set; }
        public string especificacion { get; set; }
        public string metodologia { get; set; }
        [Required]
        public string resultado { get; set; }
        public string unidadMedida { get; set; }
        public string valor { get; set; }
    }
}
