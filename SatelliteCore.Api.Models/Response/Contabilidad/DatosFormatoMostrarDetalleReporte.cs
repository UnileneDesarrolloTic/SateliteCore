namespace SatelliteCore.Api.Models.Response.Contabilidad
{
    public struct DatosFormatoMostrarDetalleReporte
    {
        public string TipoDocumento { get; set; }
        public string NumeroDocumento { get; set; }
        public string TransaccionCodigo { get; set; }
        public int Secuencia { get; set; }
        public string Item { get; set; }
        public string Lote { get; set; }
        public decimal Cantidad { get; set; }
        public decimal PrecioUnitario { get; set; }
        public decimal MontoTotal { get; set; }
        public decimal CantidadAntes { get; set; }
        public decimal PrecioUnitarioAntes { get; set; }
        public decimal MontoTotalAntes { get; set; }
    }
}
