using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request.Licitaciones
{
    public struct DatosFormatoRegistrarExpedienteLi
    {
        public string idProceso { get; set; }
        public string entrega { get; set; }
        public string item { get; set; }
        public string usuario { get; set; }
        public string ordencompra { get; set; }
        public string factura { get; set; }
        public string expediente { get; set; }
        public List<DatosFormatoDetalleExpediente> detalleExpediente { get; set; }
    }
}
