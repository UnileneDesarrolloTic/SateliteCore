using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatoListarPersonaTecnica
    {
       public int idEmpleado { get; set; }
       public string TipoPlanilla { get; set; }
       public string  NombreCompleto { get; set; }
       public int TipoHorario { get; set; }
       public string Documento { get; set; }
       public bool isSelected { get; set; }
    }
}
