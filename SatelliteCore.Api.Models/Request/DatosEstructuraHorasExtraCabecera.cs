using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request
{
    public class DatosEstructuraHorasExtraCabecera
    {   
        public int idCodigo { get; set; }
        public int Area { get; set; }
        public DateTime FechaRegistro { get; set; }
        public string Justificacion { get; set; }
        public string Persona { get; set; }
        public string Estado { get; set;  }
        public string Periodo { get; set; }
        public List<DatosEstructuraHorasExtrasDetalle> ListaPersona { get; set; }

    }
}
