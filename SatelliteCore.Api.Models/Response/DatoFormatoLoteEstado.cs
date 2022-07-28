using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatoFormatoLoteEstado
    {   
        public int Id { get; set; }
        public string Lote { get; set; }
        public string OrdenFabricacion { get; set;}
        public DateTime FechaRegistro { get; set; }
        public string Estado { get; set; }
        public string Usuario { get; set; }
    }
}
