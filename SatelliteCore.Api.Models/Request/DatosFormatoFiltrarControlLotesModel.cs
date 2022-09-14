using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request
{
    public class DatosFormatoFiltrarControlLotesModel
    {
        public string FechaInicio { get; set; }
        public string FechaFinal { get; set; }
        public string OrdenFabricacion { get; set; }
        public int Pagina { get; set; }
        public int RegistrosPorPagina { get; set; }
    }
}
