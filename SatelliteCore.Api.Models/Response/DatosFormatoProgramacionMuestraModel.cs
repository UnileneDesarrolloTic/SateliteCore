using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public class DatosFormatoProgramacionMuestraModel
    {
        public int IdProgramacion { get; set; }
        public int NumeroEntrega { get; set; }
        public int NumeroItem { get; set; }
        public string DescripcionItem { get; set; }
        public string CodItem { get; set; }
        public string NumeroMuestreo { get; set; }
        public string NumeroEnsayo { get; set; }
        public int IdProceso { get; set; }
        public string Protocolo { get; set; }

    }
}
