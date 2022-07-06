using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public class FormatoGuiaPorFacturarModel
    {
        public string SerieNumero { get; set; }
        public string GuiaNumero { get; set; }
        public int Destinatario { get; set; }
        public string DestinatarioRUC { get; set; }
        public string DestinatarioNombre { get; set; }
        public DateTime FechaDocumento { get; set; }
        public string EstadoGuia { get; set; }
        public string FacturaNumero { get; set; }
        public DateTime? FacturaFecha { get; set; }
        public string EstadoFacturacion { get; set; }
        public string DestinatarioDireccion { get; set; }
        public string ReferenciaNumeroOrden { get; set; }
        public string UltimoUsuario { get; set; }
        public string ComentarioGuia { get; set; }
        public string ReprogramacionPuntoPartida { get; set; }
        public string EstadoLogistica { get; set; }
        public string LicitacionNumeroProceso { get; set; }
        public bool ComentariosEntrega { get; set; }
        public string UsuComercial { get; set; }
    }
}
