namespace SatelliteCore.Api.Models.Response
{
    public struct ListarDetallePedido
    {
        public string Pedido { get; set; }
        public int Linea { get; set; }
        public string Item { get; set; }
        public string Descripcion { get; set; }
        public decimal CantidadPedida { get; set; }
        public decimal PrecioUnitario { get; set; }
        public decimal Monto { get; set; }
    }
}
