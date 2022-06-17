using System;

namespace SatelliteCore.Api.Models.Entities
{
    public struct ConfiguracionEntity
    {
        public int IdConfiguracion { get; set; }
        public int Id { get; set; }
        public string Grupo { get; set; }
        public string ValorTexto1 { get; set; }
        public string ValorTexto2 { get; set; }
        public int? ValorEntero1 { get; set; }
        public int? ValorEntero2 { get; set; }
        public int? ValorEntero3 { get; set; }
        public decimal? ValorDecimal1 { get; set; }
        public decimal? ValorDecimal2 { get; set; }
        public decimal? ValorDecimal3 { get; set; }
        public DateTime? ValorFecha1 { get; set; }
        public DateTime? ValorFecha2 { get; set; }
        public DateTime? ValorFecha3 { get; set; }
        public bool? ValorBit { get; set; }
        public string Estado { get; set; }
    }
}
