using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatoHorasExtrasDetalle
    {
        public int IdDetalle { get; set; }
        public int IdCabecera { get; set; }
        public int IdPersona { get; set; }
        public string Documento { get; set; }
        public string NombrePersona { get; set; }
        public decimal horasextras { get; set; }
    }
}
