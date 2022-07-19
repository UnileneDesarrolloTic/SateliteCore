using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request
{
    public  class FormatoListarMaestroItemModel
    {
        public string Item { get; set; }
        public string DescripcionLocal { get; set; }
        public string DescripcionIngles { get; set; }
        public string NumeroDeParte { get; set; }
        public string Estado { get; set; }
        public string CaracteristicaValor05 { get; set; }
    }
}
