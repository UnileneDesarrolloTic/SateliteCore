using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatoDatosProductoCostobase
    {
        public DateTime UltFechaDoc { get; set; }
        public string CodigoItem { get; set; }
        public string NumeroDeParte { get; set; }
        public string Item { get; set; }
        public string Familia { get; set; }
        public string Subfamilia { get; set; }
        public decimal CostoMateriaPrima_Soles { get; set; }
        public decimal CostoManoObra_Soles { get; set; }
        public decimal CostoIndirectoFab_Soles { get; set; }
        public decimal Mto_CostoBase { get; set; }
    }
}
