using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.ProgramacionOperaciones
{
    public struct DatosFormatoListadoFechaProgramadas
    {
        public string OrdenFabricacion { get; set; }
        public string Tipo { get; set; }
        public DateTime Fecha { get; set; }
        public string Comentario { get; set; }
        public string UsuarioCreacion { get; set; }
        public DateTime FechaRegistro { get; set; }
    }
}
