using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatoObtenerAgujasDescripcionModel
    {
        public int ID_AGUJA { get; set; }
        public string NOMENCLATURA { get; set; }
        public string DESCRIPCIONAGUJA { get; set; }
        public int ID_CARACTERISTICA { get; set; }
        public int ID_DESCRIPCION { get; set; }
        public string DESCRIPCIONLOCAL { get; set; }
        public string DESCRIPCIONINGLES { get; set; }
    }
}
