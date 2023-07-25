using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.GestioEquipoEngaste
{
    public struct DatosFormularioFiltroEquipo
    {
        public int persona { get; set; }
        public string tipo { get; set; }
        public int codigo { get; set; }
    }
}
