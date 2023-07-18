using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.Common
{
    public struct DatosFormatoInformacionItem
    {
        public string Item { get; set; }
        public string ItemTipo { get; set; }
        public string Linea { get; set; }
        public string Familia { get; set; }
        public string SubFamilia { get; set; }
        public string DescripcionLocal { get; set; }
        public string DescripcionIngles { get; set; }
        public string NumeroDeParte { get; set; }
        public string UnidadCodigo { get; set; }
    }
}
