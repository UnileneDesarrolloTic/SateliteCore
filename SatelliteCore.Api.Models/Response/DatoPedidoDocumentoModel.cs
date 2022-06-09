using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatoPedidoDocumentoModel
    {
        public string NumeroDocumento { get; set; }
        public string ClienteNombre { get; set; }
        public DateTime FechaDocumento { get; set; }
        public string RInterno { get; set; }
        public string RExterno { get; set; }
    }
}
