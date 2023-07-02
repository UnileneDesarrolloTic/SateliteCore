using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.CompraImportacion
{
    public struct DatosFormatoArimaNacionalImportada
    {
        public int Arima { get; set; }
        public string Material { get; set; }
        public string Item { get; set; }
        public string DescripcionLocal { get; set; }
        public int DiasEspera { get; set; }
        public int PuntoCorte { get; set; }
        public int Pronostico { get; set; }
        public int MaximoStock { get; set; }
        public DateTime Fechaproxima { get; set; }
        public int CantidadComprarAprox { get; set; }
    }
}
