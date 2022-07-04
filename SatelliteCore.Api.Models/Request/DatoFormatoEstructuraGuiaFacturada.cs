using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request
{
    public  class DatoFormatoEstructuraGuiaFacturada
    {
        public bool comentariosEntrega { get; set; }
        public int destinatario { get; set; }
        public string guiaNumero { get; set; }
        public string serieNumero { get; set; }
        
    }
}
