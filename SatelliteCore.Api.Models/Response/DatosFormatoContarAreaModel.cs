using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatoContarAreaModel
    {
        public int IdArea { get; set; }
        public string Descripcion { get; set; }
        public int Total { get; set; }
        public int Asistio { get; set; }
        public int falto { get; set; }
    }
}
