using System;

namespace SatelliteCore.Api.Models.Request
{
    public struct DatosProtocoloAnalisisListado
    {
        public DateTime? FechaInicio { get; set; }
        public DateTime? FechaFin { get; set; }
        public string NumeroDocumento { get; set; }
        public string Lote { get; set; }
        public string OrdenFabricacion { get; set; }
        public int IdCliente { get; set; }
        public string TipoDocumento { get; set; }
        public string Protocolo { get; set; }
    }
}
