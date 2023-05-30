namespace SatelliteCore.Api.Models.Request.GestionOrdenesServicio
{
    public struct DatosTransportistaDTO
    {
        public int Id { get; set; }
        public string Descripcion { get; set; }
        public string Direccion { get; set; }
        public string Ruc { get; set; }
        public string Telefono_1 { get; set; }
        public string Telefono_2 { get; set; }
        public string Email { get; set; }
        public string Detalle { get; set; }
    }
}
