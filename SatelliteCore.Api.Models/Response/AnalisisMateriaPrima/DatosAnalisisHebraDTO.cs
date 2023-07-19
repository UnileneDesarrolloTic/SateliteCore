using System;

namespace SatelliteCore.Api.Models.Response.AnalisisMateriaPrima
{
    public struct DatosAnalisisHebraDTO
    {
        public string Producto { get; set; }
        public string Item { get; set; }
        public string NumeroDeParte { get; set; }
        public string Proveedor { get; set; }
        public string Lote { get; set; }
        public string NroIngreso { get; set; }
        public DateTime FechaIngreso { get; set; }
        public string OrdCompra { get; set; }
        public string Analisis { get; set; }
        public decimal Longitud { get; set; }
        public decimal MinimoDiametro { get; set; }
        public decimal MaximoDiametro { get; set; }
        public decimal Tension { get; set; }
        public bool RegistroLongitud { get; set; }
    }
}
