using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request
{
    public class DatosFormatoCabeceraFormatoProtocolo
    {
        public string Idioma { get; set; }
        public DateTime fechaanalisis { get; set; }
        public string NumeroLote { get; set; }
        public string NumeroParte { get; set; }
        public string Tecnica { get; set; }
        public string Metodo { get; set; }
        public string Detalle { get; set; }
        public List<DatosFormatoDetalleFormatoProtocolo> TablaPrueba { get; set; }
    }
}
