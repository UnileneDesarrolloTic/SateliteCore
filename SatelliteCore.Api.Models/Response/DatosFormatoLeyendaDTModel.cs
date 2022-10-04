using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public class DatosFormatoLeyendaDTModel
    {
        public int IdLeyenda { get; set; }
        public string RegistroSanitario { get; set; }
        public string Marca { get; set; }
        public string Hebra { get; set; }
        public string TecnicaEspaniol { get; set; }
        public string TecnicaIngles { get; set; }
        public string MetodoEspaniol { get; set; }
        public string MetodoIngles { get; set; }
        public string DetalleEspaniol { get; set; }
        public string DetalleIngles { get; set; }
    }
}
