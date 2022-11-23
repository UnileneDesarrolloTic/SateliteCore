namespace SatelliteCore.Api.Models.Dto.GestionCalidad
{
    public struct RespuestaReclamoDTO
    {
        public int IdDetalle { get; set; }
        public string Estado { get; set; }
        public string TipoEnvio { get; set; }
        public string Destinatario { get; set; }
        public string LoteCanje { get; set; }
        public string Respuesta { get; set; }
        public string Usuario { get; set; }

        public bool ValidarRegistro()
        {
            if(IdDetalle == 0 || string.IsNullOrEmpty(Estado) || string.IsNullOrEmpty(TipoEnvio) || string.IsNullOrEmpty(Destinatario) 
                || string.IsNullOrEmpty(Respuesta) || string.IsNullOrEmpty(Usuario))
                return false;

            if (Estado != "A" && Estado != "R")
                return false;

            return true;
        }
    }
}
