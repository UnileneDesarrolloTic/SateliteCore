using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request.ComprobanteOrdenCompra
{
    public struct DatosFormatoDetalleOrdenCompraRequest
    {
        public bool seleccionar { get; set; }
        public string documento { get; set; }
        public string item { get; set; }
        public int secuencia { get; set; }
        public string descripcion { get; set; }
    }
}
