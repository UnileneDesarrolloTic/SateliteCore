using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Entities
{
    public class AlmacenamientoSsomaEntity
    {
           public int idSsomaAlmacenamiento { get; set; }
           public string Descripcion { get; set; }
           public string UsuarioCreacion { get; set; }
           public DateTime FechaModificacion { get; set; }
           public string Estado { get; set; }
    }
}
