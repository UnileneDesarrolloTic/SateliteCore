using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public class FormatoDatosBusquedaItemsVentas
    {
        public int idAgrupador { get; set; }
        public int idSubAgrupador { get; set; }
        public string idLinea { get; set; }
        public string idfamilia { get; set; }
        public string idSubFamilia { get; set; }
        public string idmarca { get; set; }
        public string Item { get; set; }
        public string Codsut { get; set; }
        public string Descripcion { get; set; }
        public int Origen { get; set; }

    }
}
