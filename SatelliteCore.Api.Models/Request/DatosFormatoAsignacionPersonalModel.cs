using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request
{
    public class DatosFormatoAsignacionPersonalModel
    {
          public int IdArea { get; set; }
          public List <PersonaLaboral> ListaPersona { get; set; }
    }

    public struct PersonaLaboral
    {
        public int idPersona { get; set; }
    }
}
