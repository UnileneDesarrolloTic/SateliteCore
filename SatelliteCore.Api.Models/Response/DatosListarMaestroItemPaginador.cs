using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public class DatosListarMaestroItemPaginador
    {

        public int idAgrupador { get; set; }
        public int idSubAgrupador { get; set; }
        public string idLinea { get; set; }
        public string Descripcion { get; set; }
        public string NumeroParte { get; set; }
        public string idSubFamilia { get; set; }
        public string idestado { get; set; }
        public string idfamilia { get; set; }
        public string idmarca { get; set; }
        public string item { get; set; }
        public int Pagina { get; set; }
        public int RegistrosPorPagina { get; set; }
    }
}
