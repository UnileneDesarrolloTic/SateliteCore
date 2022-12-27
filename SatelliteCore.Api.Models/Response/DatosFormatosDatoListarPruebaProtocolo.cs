using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
   public struct DatosFormatosDatoListarPruebaProtocolo
    {
        public int ID_PRUEBA { get; set; }
        public int ORDEN { get; set; }
        public string DescripcionLocal { get; set; }
        public string UnidadMedida { get; set; }
        public string ESPECIFICACION { get; set; }
        public string VALOR { get; set; }
        public string RESULTADO { get; set; }
        public string Metodologia { get; set; }
        public string decimales { get; set; }
        public string PreResultado { get; set; }
        public string FlagPreResultado { get; set; }
    }
}
