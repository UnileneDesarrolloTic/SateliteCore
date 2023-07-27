namespace SatelliteCore.Api.Models.Response.TransferenciaPT
{
    public struct RegistrarRecepcionPtDTO
    {
        public int IdDetalle { get; set; }
        public string ControlNumero { get; set; }
        public string Lote { get; set; }
        public string OrdenFabricacion { get; set; }
        public int UsuarioRecepcion { get; set; }
        public string ItemCodigo { get; set; }
        public decimal CantidadParcial { get; set; }
        public string AlmacenDestino{ get; set; }


        public bool ValidarDatos()
        {
            if(string.IsNullOrWhiteSpace(ControlNumero) || string.IsNullOrWhiteSpace(Lote) || string.IsNullOrWhiteSpace(OrdenFabricacion) || UsuarioRecepcion == 0
                || string.IsNullOrWhiteSpace(ItemCodigo) || string.IsNullOrWhiteSpace(AlmacenDestino) || CantidadParcial == (decimal)0.0 )
                return false;
            return true;
        }
    }
}
