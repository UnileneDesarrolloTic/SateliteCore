using System;

namespace SatelliteCore.Api.Models.Response
{
    public class ListarGuiaInformeLPModel
    {
        public string SerieNumero { get; set; }
        public string GuiaNumero { get; set; }
        public string OrdenCompra { get; set; }
        public string Estado { get; set; }
        public string Comentario { get; set; }
        public string EntregaPecosa { get; set; }
        public string EstadoLogistica { get; set; }
        public string ComentarioSalida { get; set; }
        public DateTime fechadocumento { get; set; }

    }
}
