using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.GestioEquipoEngaste
{
    public struct DatosFormatoDetalleInformacionEquipo
    {
        public bool Seleccionar { get; set; }
        public int Id { get; set; }
        public string Codigo { get; set; }
        public string Dado { get; set; }
        public string Alambre { get; set; }
        public string Tipo { get; set; }
    }
}
