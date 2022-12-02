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
        public string ItemComponente { get; set; }
        public string LineaItemComponente { get; set; }
        public string FamiliaItemComponente { get; set; }
        public string SubFamiliaItemComponente { get; set; }
        public decimal Cantidad { get; set; }
        public decimal CostoUnitarioSoles { get; set; }
        public decimal CostoUnitarioDolares { get; set; }
        public decimal CostoUnitario { get; set; }
    }
}
