using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.ComprobanteOrdenCompra
{
    public struct DatosFormatoDetalleOrdenCompra
    {   public bool Seleccionar { get; set; }
        public string Documento { get; set; }
        public string Item { get; set; }
        public int Secuencia { get; set; }
        public string Descripcion { get; set; }
    }
}
