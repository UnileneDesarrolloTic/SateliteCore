using SatelliteCore.Api.Models.Response;
using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request
{
    public class DatosFormatoCabeceraOrdenCompraModel
    {
        public string Documento { get; set; }
        public string FechaLlegada { get; set; }
        public List<DatosFormatoItemActualizarItemOrdenCompra> Detalle { get; set; }
    }
}
