using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct FormatoDatoInformacionItem
    {
        public string Item { get; set;  }
        public string Descripcion { get; set; }
        public decimal Aduanas { get; set; }
        public decimal Almacen { get; set; }
    }
}
