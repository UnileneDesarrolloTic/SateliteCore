using System;


namespace SatelliteCore.Api.Models.Request
{
    public class DatosFormatoFiltraHoraExtras
    {
        public DateTime FechaInicio { get; set; }
        public DateTime FechaFin { get; set; }
        public string Estado { get; set; }
        public string Periodo { get; set; }
        public bool TipoFiltro { get; set; }
    }
}
