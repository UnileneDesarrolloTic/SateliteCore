namespace SatelliteCore.Api.Models.Response.GestionCalidad
{
    public struct DatosLoteReclamoDTO
    {
        public string Item { get; set; }
        public string DescripcionLocal { get; set; }
        public string Linea { get; set; }
        public string Familia { get; set; }
        public string SubFamilia { get; set; }
        public string Marca { get; set; }
    }
}
