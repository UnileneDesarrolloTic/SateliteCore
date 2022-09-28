using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatoListarTransaccion
    {
        public string Lote { get; set; }
        public string Periodo { get; set; }
        public string Documento { get; set; }
        public decimal Cantidad { get; set; }
        public string AlmacenCodigo { get; set; }
        public string DocumentoTransaccion { get; set; }
    }
}
