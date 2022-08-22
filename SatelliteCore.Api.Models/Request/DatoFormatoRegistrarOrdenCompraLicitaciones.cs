using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request
{
    public class DatoFormatoRegistrarOrdenCompraLicitaciones
    {
        public int idProceso { get; set; }
        public int NumeroEntrega { get; set; }
        public int Item { get; set; }
        public string Usuario { get; set; }
        public string NumeroOC { get; set; }
        public int CantidadOC { get; set; }
    }
}
