using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.CompraImportacion
{
    public class DatosFormatoMostrarOrdenCompraNacionalImportacion
    {
        public string Item { get; set; }
        public string NumeroOrden { get; set; }
        public string Proveedor { get; set; }
        public decimal Cantidad { get; set; }
        public decimal CantidadRecibida { get; set; }
        public decimal PendienteOC { get; set; }
        public DateTime FechaPreparacion { get; set; }
        public DateTime Fecha { get; set; }
        public int DiferenciaFecha { get; set; }
        public DateTime FechaLlegada { get; set; }
    }
}
