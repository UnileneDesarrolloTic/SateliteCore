using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Request.Contabildad
{
    public class DatosFormatoRestablecerCierre
    {
        public string Tipo { get; set; }
        public string Periodo { get; set; }
        public List<DatosFormatoEstablecerDetalleCierre> Detalle { get; set; }
    }
}
