using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct FormatoDatosOrdenCompraItem
    {
        public string NumeroOrden { get; set; }
        public string NombreCompleto { get; set; }
        public string Documento { get; set; }
        public decimal CantidadPedida { get; set; }
        public DateTime FechaPrometida { get; set; }
        public string Item { get; set; }
    }
}
