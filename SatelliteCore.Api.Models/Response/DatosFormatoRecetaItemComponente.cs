using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatoRecetaItemComponente
    {
        public string Periodo { get; set; }
        public string Item { get; set; }
        public string NombreProducto { get; set; }
        public string NumeroDeParte { get; set; }
        public decimal SECUENCIA { get; set; }
        public string ITEMCOMPONENTE { get; set; }
        public decimal Cantidad { get; set; }
        public decimal CostoUnitarioSoles { get; set; }
        public string MATERIAL { get; set; }
        public string Linea { get; set; }
        public string Familia { get; set; }
        public string SubFamilia { get; set; }

    }
}
