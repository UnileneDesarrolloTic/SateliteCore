using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request
{
    public class DatosRequestFormatoContratoProcesoModel
    {
        public int idproceso { get; set; }
        public string tipodeusuario { get; set; }
        public int numeroitem { get; set; }
        public string numeroContrato { get; set; }
    }
}
