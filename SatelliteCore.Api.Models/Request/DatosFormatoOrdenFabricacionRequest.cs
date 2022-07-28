using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request
{
    public class DatosFormatoOrdenFabricacionRequest
    {
        public string cliente { get; set; }
        public  int contraMuestra { get; set; }
        public string descripcionLocal { get; set; }
        public string fechaProduccion { get; set; }
        public string item { get; set; }
        public string lote { get; set; }
        public string marca { get; set; }
        public string numeroCaja { get; set; }
        public string numeroParte { get; set; }

    }
}
