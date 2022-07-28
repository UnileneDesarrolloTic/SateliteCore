using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct FormatoEstructuraLoteEtiquetas
    {
        public DateTime FechaProduccion { get; set; }
        public string Item { get; set; }
        public string NumeroParte { get; set; }
        public string Marca { get; set; }
        public string DescripcionLocal { get; set; }
        public string Cliente { get; set; }
        public string Lote { get; set; }
        public string OrdenFabricacion { get; set; }
        public string transferidoflag { get; set; }

    }
}
