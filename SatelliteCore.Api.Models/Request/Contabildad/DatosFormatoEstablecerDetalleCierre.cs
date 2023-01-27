using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request.Contabildad
{
    public struct DatosFormatoEstablecerDetalleCierre
    {
        public string tipoDocumento { get; set; }
        public string numeroDocumento { get; set; }
        public string transaccionCodigo { get; set; }
        public int secuencia { get; set; }
        public string item { get; set; }
        public string lote { get; set; }
        public decimal cantidad { get; set; }
        public decimal precioUnitario { get; set; }
        public decimal montoTotal { get; set; }
        public decimal cantidadAntes { get; set; }
        public decimal precioUnitarioAntes { get; set; }
        public decimal montoTotalAntes { get; set; }
    }
}
