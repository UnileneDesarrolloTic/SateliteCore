using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatoLenguajeFormatoProtocolo
    {
        public string NvcTitulo { get; set; }
        public string NvcProducto { get; set; }
        public string NvcPresentacion { get; set; }
        public string NvcFechaFabricacion { get; set; }
        public string NvcLote { get; set; }
        public string NvcTamanioLote { get; set; }
        public string NvcFechaExpira { get; set; }
        public string NvcMarca { get; set; }
        public string NvcFechaAnalisis { get; set; }
        public string NvcObservacion { get; set; }
        public string NvDPruebasEfectuadas { get; set; }
        public string NvDEspecificaciones { get; set; }
        public string NvDResultados { get; set; }
        public string NvDMetodologia { get; set; }
    }
}
