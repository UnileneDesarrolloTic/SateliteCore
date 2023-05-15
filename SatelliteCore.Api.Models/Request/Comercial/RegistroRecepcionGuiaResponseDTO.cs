using System;

namespace SatelliteCore.Api.Models.Request.Comercial
{
    public struct RegistroRecepcionGuiaResponseDTO
    {
        public string GuiaNumero { get; set; }
        public string Destinatario { get; set; }
        public string Factura { get; set; }
        public DateTime FechaDocumento { get; set; }
    }
}
