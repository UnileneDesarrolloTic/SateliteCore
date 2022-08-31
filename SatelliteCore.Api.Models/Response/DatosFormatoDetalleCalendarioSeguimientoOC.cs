using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatoDetalleCalendarioSeguimientoOC
    {
        public string Fecha { get; set; }
        public string Clasificacion { get; set; }
        public string Item { get; set; }
        public string DescripcionLocal { get; set; }
        public string NumeroOrden { get; set; }
        public string DescripcionCompleta { get; set; }
        public decimal Cantidad { get; set; }
    }
}
