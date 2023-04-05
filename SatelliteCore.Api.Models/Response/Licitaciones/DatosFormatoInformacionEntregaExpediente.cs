using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.Licitaciones
{
   public struct DatosFormatoInformacionEntregaExpediente
    {   
        public string Codigo { get; set; }
        public string Destino { get; set; }
        public DateTime? FechaEntrega { get; set; }
        public DateTime? FechaAprobacion { get; set; }
        public string Usuario { get; set; }
    }
}
