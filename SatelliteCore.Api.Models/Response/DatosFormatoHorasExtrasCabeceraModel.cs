using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatoHorasExtrasCabeceraModel
    {
        public int IdCabecera { get; set; }
        public int IdArea { get; set; }
        public string Descripcion { get; set; }
        public string Justificacion { get; set; }
        public DateTime FechaRegistro { get; set; }
        public string TipoPersona { get; set; }
        public string Estado { get; set; }
        public string Periodo { get; set; }
    }
}
