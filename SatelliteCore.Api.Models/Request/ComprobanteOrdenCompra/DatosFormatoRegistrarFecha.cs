using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request.ComprobanteOrdenCompra
{
    public struct DatosFormatoRegistrarFecha
    {   
        public string ordenCompra { get; set; }
        public string comentario { get; set; }
        public DateTime prometida { get; set; }
        public List<DatosFormatoDetalleOrdenCompraRequest> detalle { get; set; }
    }
}
