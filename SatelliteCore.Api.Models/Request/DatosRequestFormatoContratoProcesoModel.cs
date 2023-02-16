using System.ComponentModel.DataAnnotations;

namespace SatelliteCore.Api.Models.Request
{
    public class DatosRequestFormatoContratoProcesoModel
    {   [Required]
        public int idproceso { get; set; }
        [Required]
        public string tipodeusuario { get; set; }
        [Required]
        public int numeroitem { get; set; }
        public string numeroContrato { get; set; }
    }
}
