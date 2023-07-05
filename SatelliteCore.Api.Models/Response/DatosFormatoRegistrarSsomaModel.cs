using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public class DatosFormatoRegistrarSsomaModel
    {
        public int idSsoma { get; set; }
        public string codigo { get; set; }
        public string nombreDocumento { get; set; }
        public int tipoDocumento { get; set; }
        public int version { get; set; }
        public  int vigencia {get; set; }
        public string fechapublicacion { get; set; }
        public string fecharevision { get; set; }
        public string fechaAprobacion { get; set; }
        public int estado { get; set; }
        public int Ubicacion { get; set; }
        public int Almacenamiento { get; set; }
        public int proteccion { get; set; }
        public string responsable { get; set; }
        public int archivopasivo { get; set; }
        public string comentario { get; set; }
    }
}
