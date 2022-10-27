using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatoComponentePrecioUnitario
    {
        public string ItemComponente { get; set; }
        public string DescripcionLocal { get; set; }
        public string Linea { get; set; }
        public string Familia { get; set; }
        public string SubFamilia { get; set; }
        public decimal CostoUnitarioSoles { get; set; }
        public decimal CostoUnitarioDolares { get; set; }
        public string NombreCompleto { get; set; }
    }
}
