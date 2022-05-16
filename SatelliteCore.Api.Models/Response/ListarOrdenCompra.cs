namespace SatelliteCore.Api.Models.Response
{
    public struct ListarOrdenCompra
    {
        public int Secuencia { get; set; }
        public string ControlNumero { get; set; }
        public string Item { get; set; }
        public string DescripcionItem { get; set; }
        public string UnidadCodigo { get; set; }
        public int CantidadPedida { get; set; }
        public int CantidadRecibida { get; set; }
        public string LoteAprobado { get; set; }
        public string LoteRechazado { get; set; }
        public int CodProveedor { get; set; }
        public string NumeroOrden { get; set; }
    }
}
