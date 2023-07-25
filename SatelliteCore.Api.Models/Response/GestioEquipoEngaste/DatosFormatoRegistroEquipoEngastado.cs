using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.GestioEquipoEngaste
{
    public struct DatosFormatoRegistroEquipoEngastado
    {
        public int idEquipo { get; set; }
        public string nombre { get; set; }
        public int idpersona { get; set; }
        public string persona { get; set; }
        public string Tipo { get; set; }
        public string estado { get; set; }
        public List<int> detalle { get; set; }

    }
}
