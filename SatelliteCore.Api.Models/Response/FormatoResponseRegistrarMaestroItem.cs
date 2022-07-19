using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct FormatoResponseRegistrarMaestroItem
    {
        public string Item { get; set; }
        public string DescripcionLocal { get; set; }
        public string DescripcionCompleta { get; set; }
        public string Estado { get; set; }
        public string Comentario { get; set; }
        public bool Respuesta { get; set; }

    }
}
