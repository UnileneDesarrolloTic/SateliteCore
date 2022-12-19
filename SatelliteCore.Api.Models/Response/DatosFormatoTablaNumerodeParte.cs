using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatoTablaNumerodeParte
    {
        public string Grupo { get; set; }
        public string NombreGrupo { get; set; }
        public string CodigoTabla { get; set; }
        public string NombreTabla { get; set; }
        public string Codigo { get; set; }
        public string DescripcionLocal { get; set; }
        public int Longitud { get; set; }
        public string Estado { get; set; }
        public DateTime UltimaFechaModif { get; set; }
        public string UltimoUsuario { get; set; }
    }
}
