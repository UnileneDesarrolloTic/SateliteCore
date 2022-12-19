using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatoObtenerTablaAgujasNuevoModel
    {
        public int ID_AGUJA { get; set; }
        public string NOMENCLATURA { get; set; }
        public string DEPENDENCIA { get; set; }
        public string DESCRIPCIONLOCAL { get; set; }
        public string DESCRIPCIONINGLES { get; set; }
        public string ESTADO { get; set; }
        public int CANT_AGUJA { get; set; }
    }
}
