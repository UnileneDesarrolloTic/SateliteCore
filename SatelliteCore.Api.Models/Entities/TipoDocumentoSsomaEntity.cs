using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Entities
{
    public class TipoDocumentoSsomaEntity
    {
        public int idTipoDocumento { get; set; }
        public string Descripcion { get; set; }
        public string Codigo { get; set; }
        public string UsuarioCreacion { get; set; }
        public DateTime FechaModificacion { get; set; }
        public string Estado { get; set; }
    }
}
