using System.ComponentModel.DataAnnotations;

namespace SatelliteCore.Api.Models.Request
{
    public struct ActualizarReporteCotizacionModel
    {
        [Required(ErrorMessage = "El identificador del reporte es obligatorio")]
        public string IdObject { get; set; }

        [Required(ErrorMessage = "Los datos de la cotización obligatorio")]
        public object Cotizacion { get; set; }
    }
}
