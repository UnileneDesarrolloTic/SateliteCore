using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request
{
    public class DatosEstructuraHorasExtrasDetalle
    {
        public int persona { get; set; }
        public string nombreCompleto { get; set; }
        public decimal horasextras { get; set; }
    }
}
