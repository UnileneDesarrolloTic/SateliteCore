using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request
{
    public class DatosFormatoFiltrarAnalisisCostoRequest
    {
        public string CodProducto { get; set; }
        public string NumeroCotizacion { get; set; }
        public Boolean Opcion { get; set; }
        public string base64 { get; set; }
        public Boolean BusquedaExcel { get; set; }
        public string idSubFamilia { get; set; }
        public string idfamilia { get; set; }
    }
}
