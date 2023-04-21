using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatoInformacionItemOrdenCompra
    {
        public List<FormatoDatosOrdenCompraItem> ListaOrdenCompra { get; set; }
        public List<object> Detalle { get; set; }
    }
}
