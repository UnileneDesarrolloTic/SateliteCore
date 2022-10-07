using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatoDetalleRecetaMPLogistica
    {
        public string Item { get; set;  }
        public decimal Secuencia { get; set; }
        public string ItemComponente { get; set; }
        public decimal CantidadNeta { get; set; }
        public string ItemTipo { get; set; }
        public string DescripcionLocal { get; set; }
        public string UnidadCodigo { get; set; }
        public decimal StockActual { get; set; }
        public decimal StockDisponible { get; set; }
        public string ItemAlternativo { get; set; }
    }
}
