namespace SatelliteCore.Api.Models.Request
{
    public struct ListarAnalisisAgujasFiltroPg
    {
        public string OrdenCompra { get; set; }
        public string Lote { get; set; }
        public int RegistroPorPagina { get; set; }
        public int Pagina { get; set; }
    }
}
