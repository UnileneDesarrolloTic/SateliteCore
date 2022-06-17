using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request
{
    public class DatosFormatoDetalleProceso
    {
        public string Pedido { get; set; }
        public int Linea { get; set; }
        public string Item { get; set; }
        public string Descripcion { get; set; }
        public decimal CantidadPedida { get; set; }
        public decimal PrecioUnitario { get; set; }
        public decimal Monto { get; set; }
    }
}
