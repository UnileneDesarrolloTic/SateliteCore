using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.OCDrogueria
{
    public struct DatosFormatoProyeccionDrogueria
    {
        public string Item { get; set; }
        public string Descripcion { get; set; }
        public int DiasEspera { get; set; }
        public int PuntoCorteDebePagar { get; set; }
        public int Pronostico { get; set; }
        public int MaximoStock { get; set; }
        public DateTime FechaProxima { get; set; }
        public int CantidadComprarAprox { get; set; }
    }
}
