using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatoListarCotizacionExportacion
    {
        public string CompaniaSocio { get; set; }
        public string NumeroDocumento { get; set; }
        public DateTime FechaDocumento { get; set; }
        public int ClienteNumero { get; set; }
        public string ClienteNombre { get; set; }
        public string MonedaDocumento { get; set; }
        public decimal MontoTotal { get; set; }
        public string Comentarios { get; set; }
        public string Estado { get; set; }
        public int Vendedor { get; set; }
        public decimal MontoAfecto { get; set; }
        public decimal MontoNoAfecto { get; set; }
        public decimal MontoDescuentos { get; set; }
        public string PreparadoPor_Nombre { get; set; }
        public string Sucursal { get; set; }
        public string TipoCliente { get; set; }
        public string TipoVenta { get; set; }
        public string DescripcionEstado { get; set; }
        public string DescripcionMoneda { get; set; }

    }
}
