using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public class DatosFormatoFormularioCotizacionExportaciones
    {
        public int Cliente { get; set; }
        public string clienteNombre { get; set; }
        public string NumeroDocumento { get; set; }
        public string Comentarios { get; set; }
        public string Contacto { get; set; }
        public string LugarEntrega { get; set; }
        public bool FormularioNuevo { get; set; }
        public decimal MontoAfecto { get; set; }
        public decimal MontoNoAfecto { get; set; }
        public decimal Descuento { get; set; }
        public decimal ImpVentas { get; set; }
        public decimal AjusteRedondeo { get; set; }
        public decimal MontoTotal { get; set; }
        public List<FormatoDetalleCotizacionExportaciones> DetalleCotizacion { get; set; }


    }
}
