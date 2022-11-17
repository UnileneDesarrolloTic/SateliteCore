namespace SatelliteCore.Api.Models.Response
{
    public struct ObtenerAnalisisAgujaModel
    {
        public string ControlNumero { get; set; }
        public string OrdenCompra { get; set; }
        public string Item { get; set; }
        public string DescripcionItem { get; set; }
        public int CodProveedor { get; set; }
        public string Proveedor { get; set; }
        public int CantidadPruebas { get; set; }
        public string Serie { get; set; }
        public string Especialidad { get; set; }
    }
}
