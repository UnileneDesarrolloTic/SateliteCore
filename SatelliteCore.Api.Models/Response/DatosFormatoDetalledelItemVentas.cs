using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatoDetalledelItemVentas
    {
        public string Item { get; set; }
        public string AlmacenCodigo { get; set; }
        public string DescripcionLocal { get; set; }
        public string Lote { get; set; }
        public decimal StockActual { get; set; }
        public decimal StockComprometido { get; set; }
        public decimal StockDisponible { get; set; }

    }
}
