using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatoItemVentas
    {
        public string Item { get; set; }
        public string DescripcionLocal { get; set; }
        public string UnidadCodigo { get; set; }
        public string NumeroDeParte { get; set; }
        public string SubFamilia { get; set; }
        public string NombreMarca { get; set; }
        public string Presentacion { get; set; }
        public decimal StockActual { get; set; }
        public decimal StockComprometido { get; set; }
        public decimal StockDisponible { get; set; }
    }
}
