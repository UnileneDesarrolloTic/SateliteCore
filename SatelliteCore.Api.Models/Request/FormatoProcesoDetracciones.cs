using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request
{

    public class FormatoProcesoDetracciones
    {   
        public string periodo { get; set; }
        public int totalimporte { get; set; }
        public List<ProcesoDetraccionesModel> proceso { get; set; }
    }
    public class ProcesoDetraccionesModel
    {
        public int Num_Registro { get; set; }
        public int TipoDocumento { get; set; }
        public string Ruc { get; set; }
        public string RazonSocial { get; set; }
        public string NumProf { get; set; }
        public string BienServicio { get; set; }
        public int Importe { get; set; }
        public string TipoOperacion { get; set; }
        public string Periodo { get; set; }
        public string Tipo { get; set; }
        public string Serie { get; set; }
        public string Numero { get; set; }
    }
}
