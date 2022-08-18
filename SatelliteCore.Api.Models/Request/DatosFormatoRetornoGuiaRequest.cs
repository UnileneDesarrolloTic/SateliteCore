using System;
using System.ComponentModel.DataAnnotations;

namespace SatelliteCore.Api.Models.Request
{
    public struct DatosFormatoRetornoGuiaRequest
    {
        [Required]
        public string cliente { get; set; }
        
        public string fechaDocumento { get; set; }
        
        public DateTime fechaRetorno { get; set; }
        
        public string numeroGuia { get; set; }
        
        public string ordenServicios { get; set; }
    }
}
