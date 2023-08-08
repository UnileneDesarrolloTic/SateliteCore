using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.ComprobanteOrdenCompra
{
    public struct MostrarFechaPrometida
    {
        public string NumeroDocumento { get; set; }
        public int Secuencia { get; set; }
        public string Item { get; set; }
        public string Comentarios { get; set; }
        public string Reprogramacion { get; set; }
        public string Usuario { get; set; }
    }
}
