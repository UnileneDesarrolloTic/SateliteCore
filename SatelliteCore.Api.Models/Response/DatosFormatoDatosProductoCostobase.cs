using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatoDatosProductoCostobase
    {
        public string CodigoItem { get; set; }
        public string DescripcionLocal { get; set; }
        public string NumeroDeParte { get; set; }
        public string Linea { get; set; }
        public string Familia { get; set; }
        public string Subfamilia { get; set; }
        public string CodMoneda { get; set; }
        public decimal CostoMateriaPrima { get; set; }
        public decimal CostoManoObra { get; set; }
        public decimal CostoCIF { get; set; }
        public decimal CostoUnitarioBase { get; set; }
        public decimal PrecioVentaMinimo { get; set;  }
        public DateTime UltFechaDoc { get; set; }
        public string ItemComponente { get; set; }
        public string DescripcionLocalItemComponente { get; set; }
        public decimal ItemComponenteCostoDolares { get; set; }
        public decimal ItemComponenteCostUnitario { get; set; } 
        public string LineaItemComponente { get; set; }
        public string FamiliaItemComponente { get; set; }
        public string SubFamiliaItemComponente { get; set; }
        public string TipoAguja { get; set; }
        public bool isSelected { get; set; }
    }
}
