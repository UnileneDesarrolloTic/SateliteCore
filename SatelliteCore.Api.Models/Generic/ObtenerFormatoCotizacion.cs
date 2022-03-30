using System.ComponentModel.DataAnnotations;

namespace SatelliteCore.Api.Models.Generic
{
    public struct ObtenerFormatoCotizacion
    {        
        public int IdFormato { get; set; }

        [Required(ErrorMessage = "El número de cotización es obligatorio")]
        public string NroCotizacion { get; set; }

        [Required(ErrorMessage = "Los datos de la cotización obligatorio")]
        public object Cotizacion { get; set; }
    }
}
