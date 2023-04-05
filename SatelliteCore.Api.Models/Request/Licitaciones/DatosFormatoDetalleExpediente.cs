using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request.Licitaciones
{
    public struct DatosFormatoDetalleExpediente
    {
        public string codigo { get; set; }
        public string destino { get; set; }
        public DateTime? fechaEntrega { get; set; }
        public DateTime? fechaAprobacion { get; set; }
        public string usuario { get; set; }
    }
}
