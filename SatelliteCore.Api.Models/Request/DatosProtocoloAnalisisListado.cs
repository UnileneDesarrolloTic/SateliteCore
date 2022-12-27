using System.ComponentModel.DataAnnotations;

namespace SatelliteCore.Api.Models.Request
{
    public struct DatosProtocoloAnalisisListado
    {
        public string FechaInicio { get; set; }
        public string FechaFin { get; set; }
        public string NumeroDocumento { get; set; }
        public string Lote { get; set; }
        public string OrdenFabricacion { get; set; }
        public int IdCliente { get; set; }
        public string TipoDocumento { get; set; }
    }
}
