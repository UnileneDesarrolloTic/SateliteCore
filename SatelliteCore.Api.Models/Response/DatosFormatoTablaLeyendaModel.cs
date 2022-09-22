using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
   public struct DatosFormatoTablaLeyendaModel
    {
        public int ID_LEYENDA { get; set; }
        public string NUM_REGISTRO { get; set; }
        public string ID_MARCA { get; set; }
        public string ID_HEBRA { get; set; }
        public string TECNICA { get; set; }
        public string METODO { get; set; }
        public string DETALLE { get; set; }
        public string TECNICA_INGLES { get; set; }
        public string METODO_INGLES { get; set; }
        public string ESTADO { get; set; }
        public DateTime FECHA_REGISTRO { get; set; }
        public string USUARIO { get; set; }
        public string DESCRIPCIONMARCA { get; set; }
        public string DESCRIPCIONHEBRA { get; set; }
    }
}
