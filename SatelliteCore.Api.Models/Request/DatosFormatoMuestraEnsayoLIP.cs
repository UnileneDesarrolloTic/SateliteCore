using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request
{
    public class DatosFormatoMuestraEnsayoLIP
    {
        public int idProgramacion { get; set; }
        public int idProceso { get; set; }
        public int numeroItem { get; set; }
        public string numeroEnsayo { get; set; }
        public string numeroMuestreo { get; set; }
        public string presentacion { get; set; }
        public string protocolo { get; set; }
        public string temperatura { get; set; }
        public string registroSanitario { get; set; }

    }
}
