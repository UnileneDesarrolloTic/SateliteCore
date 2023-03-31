using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Entities
{
    public struct TransportistaEntity
    {
        public int Id { get; set; }
        public string Descripcion { get; set; }
        public string Direccion { get; set; }
        public string PrimerTelefono { get; set; }
        public string SegundoTelefono { get; set; }
        public string Estado { get; set; }
    }
}
