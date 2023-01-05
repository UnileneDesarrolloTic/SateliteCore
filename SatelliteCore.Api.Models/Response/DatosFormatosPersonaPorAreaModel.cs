using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatosPersonaPorAreaModel
    {
        public int Persona { get; set; }
        public string NombreCompleto { get; set; }
        public string Documento { get; set; }
    }
}
