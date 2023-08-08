using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request.ComprobanteOrdenCompra
{
    public struct DatosFormatoRegistrarFecha
    {
        public string documento { get; set; }
        public string item { get; set; }
        public int secuencia { get; set; }
        public string comentario { get; set; }
        public DateTime prometida { get; set; }
    }
}
