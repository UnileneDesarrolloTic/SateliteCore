using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatoItemActualizarItemOrdenCompra
    {
        public string numeroOrden { get; set; }
        public string nombreCompleto { get; set; }
        public string documento { get; set; }
        public decimal cantidadPedida { get; set; }
        public DateTime fechaPrometida { get; set; }
        public string item { get; set; }
    }
}
