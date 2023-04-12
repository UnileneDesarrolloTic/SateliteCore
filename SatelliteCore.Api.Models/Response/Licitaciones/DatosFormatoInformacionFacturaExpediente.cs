using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.Licitaciones
{
    public struct DatosFormatoInformacionFacturaExpediente
    {   
        public DatosFormatoFacturaProceso InformacionFactura { get; set; }
        public List<DatosFormatoInformacionEntregaExpediente> DetalleExpediente { get; set; }
        public DatosFormatoExpediente InformacionExpendiente { get; set; }
    }
}
