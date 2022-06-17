namespace SatelliteCore.Api.Models.Request
{
    public struct DatosFormatoDocumentoLicitacion
    {
        public int idcliente { get; set; }
        public string cliente { get; set; }
        public string fechainicio { get; set; }
        public string fechafinal { get; set; }
    }


}
