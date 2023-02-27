using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.OCDrogueria
{
    public class FormatosDatosDetalleOrdenCompraPrevio
    {
        public string NumeroCodigo { get; set; }
        public int Secuencia { get; set; }
        public string Item { get; set; }
        public string Descripcion { get; set; }
        public string Presentacion { get; set; }
        public int Cantidad { get; set; }
        public decimal CantidadPedida { get; set; }
        public decimal PrecioUnitario { get; set; }
        public decimal MontoTotal { get; set; }
        public string Moneda { get; set; }
        public string Estado { get; set; }
    }
}
