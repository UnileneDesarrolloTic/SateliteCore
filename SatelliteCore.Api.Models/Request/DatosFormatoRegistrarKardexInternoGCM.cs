using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request
{
    public class DatosFormatoRegistrarKardexInternoGCM
    {
        public string Lote { get; set; }
        public string OrdenFabricacion { get; set; }
        public string Transaccion { get; set; }
        public decimal Cantidad { get; set; }
        public string Comentario { get; set; }
    }
}
