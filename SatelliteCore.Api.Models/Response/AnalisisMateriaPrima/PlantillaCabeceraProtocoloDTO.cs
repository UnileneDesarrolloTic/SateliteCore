using System;

namespace SatelliteCore.Api.Models.Response.AnalisisMateriaPrima
{
    public struct PlantillaCabeceraProtocoloDTO
    {
        public string Analisis { get; set; }
        public string Item { get; set; }
        public string Proveedor { get; set; }
        public decimal Cantidad { get; set; }
        public string Unidad { get; set; }
        public string NroIngreso { get; set; }
        public string Lote { get; set; }
        public DateTime FechaVencimiento { get; set; }
        public DateTime FechaReAnalisis { get; set; }
        public DateTime FechaAnalisis { get; set; }
    }
}
