using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request
{
    public class DatosFormatoDetalleFormatoProtocolo
    {
        public int orden { get; set; }
        public string descripcionLocal { get; set; }
        public string especificacion { get; set; }
        public string metodologia { get; set; }
        public string resultado { get; set; }
        public string unidadMedida { get; set; }
        public string valor { get; set; }
    }
}
