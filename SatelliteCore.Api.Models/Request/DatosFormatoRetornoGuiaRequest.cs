using System;
using System.ComponentModel.DataAnnotations;

namespace SatelliteCore.Api.Models.Request
{
    public struct DatosFormatoRetornoGuiaRequest
    {
        [Required]
        public string cliente { get; set; }
        [Required]
        public string fechaDocumento { get; set; }
        [Required]
        public DateTime fechaRetorno { get; set; }
        [Required]
        public string numeroGuia { get; set; }
        [Required]
        public string ordenServicios { get; set; }
    }
}
