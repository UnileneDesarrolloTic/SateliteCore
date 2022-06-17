using System;

namespace SatelliteCore.Api.Models.Entities
{
    public struct FormatoLicitaciones
    {
        public DateTime FechaDocumento { get; set; }
        public int Destinatario { get; set; }
        public string SerieNumero { get; set; }
        public string GuiaNumero { get; set; }
        public string DestinatarioDireccion { get; set; }
        public int DestinatarioDireccionSecuencia { get; set; }
        public string FacturaNumero { get; set; }
        public string AlmacenCodigo { get; set; }
        public string ReferenciaNumeroPedido { get; set; }
        public string Comentarios { get; set; }

    }
}
