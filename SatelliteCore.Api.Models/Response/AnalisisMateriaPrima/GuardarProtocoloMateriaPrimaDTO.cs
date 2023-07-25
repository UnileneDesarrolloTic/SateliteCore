namespace SatelliteCore.Api.Models.Response.AnalisisMateriaPrima
{
    public struct GuardarProtocoloMateriaPrimaDTO
    {
        public string Item { get; set; }
        public string NumeroLote { get; set; }
        public string ItemNumeroDeParte { get; set; }
        public int Secuencia { get; set; }
        public string Prueba { get; set; }
        public string Especificacion { get; set; }
        public string Metodologia { get; set; }
        public string Valor { get; set; }
        public string TipoDato { get; set; }
        public decimal Minimo { get; set; }
        public decimal Maximo { get; set; }
        public string Rechazado { get; set; }
        public string Aprobado { get; set; }
        public string Conclusion { get; set; }
        public string Observaciones { get; set; }
        public string Usuario { get; set; }
    }
}
