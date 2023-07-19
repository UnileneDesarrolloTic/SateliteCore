using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.Dispensacion
{
    public class DatosFormatoInformacionDispensacionPT
    {
        public string OrdenFabricacion { get; set; }
        public string Item { get; set; }
        public string Descripcion { get; set; }
        public decimal CantidadTotal { get; set; }
        public decimal CantidadParcial { get; set; }
    }
}
