using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request
{
    public struct DatosFormatoListarSsomaModel
    {
        public int IdSsoma { get; set; }
        public int IdTipoDocumento { get; set; }
        public string TipoDocumentoDescripcion { get; set; }
        public int IdUbicacionSsoma { get; set; }
        public int IdProteccionSsoma { get; set; }
        public int IdEstadoSsoma { get; set; }
        public string EstadoDescripcion { get; set; }
        public string CodigoDocumento { get; set; }
        public string NombreDocumento { get; set; }
        public DateTime FechaPublicacion { get; set; }
        public int VersionSsoma { get; set; }
        public int Vigencia { get; set; }
        public int IdSsomaAlmacenamiento { get; set; }
        public int ArchivoPasivo { get; set; }
        public string Responsable { get; set; }
        public DateTime FechaAprobacion { get; set; }
        public DateTime FechaRevision { get; set; }
        public string Comentario { get; set; }
        public int Dias { get; set; }

    }
}
