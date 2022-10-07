using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatoMateriaPrimaItemLogistica
    {
        public string NumeroCotizacion { get; set; }
        public DateTime FechaDocumento { get; set; }
        public string clientenombre { get; set; }
        public string EstadoCotizacion { get; set; }
        public string Familia { get; set; }
        public string unidadcodigo { get; set; }
        public string ItemCodigo { get; set; }
        public string descripcion { get; set; }
        public decimal cantidadpedida { get; set; }
        public string EstadoDetalle { get; set; }
        public decimal MP { get; set; }
        public string RECETA { get; set; }
        public string Presentacion { get; set; }
    }
}
