using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request
{
    public class DatosFormatoControlLotesActualizarFEntrega
    {
        public string ordenFabricacion { get; set; }
        public string lote { get; set; }
        public int destruible { get; set; }
        public string comentarios { get; set; }
        public DateTime fechaEntrega  { get; set; }
    }
}
