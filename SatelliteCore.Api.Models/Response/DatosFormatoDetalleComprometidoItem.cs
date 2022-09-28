using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatoDetalleComprometidoItem
    {
        public string CompaniaSocio { get; set; }
        public string TipoDocumento { get; set; }
        public string NumeroDocumento { get; set; }
        public string ClienteNombre { get; set; }
        public DateTime FechaDocumento { get; set; }
        public DateTime FechaVencimiento { get; set; }
        public string TipoVenta { get; set; }
        public string Vendedor { get; set; }
        public string comentarios { get; set; }
        public string TipoDetalle { get; set; }
        public string Descripcion { get; set; }
        public string UnidadCodigo { get; set; }
        public string AlmacenCodigo { get; set; }
        public string Busqueda { get; set; }
        public string DescripcionLocal { get; set; }
        public decimal CantidadPedida { get; set; }
    }
}
