using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.Contabilidad
{
    public class DatosFormatoListadoReporteTransaccionKardex
    {
        public string Id { get; set; }
        public string Periodo { get; set; }
        public string Tipo { get; set; }
        public decimal MontoAntes { get; set; }
        public decimal MontoActual { get; set; }
    }
}
