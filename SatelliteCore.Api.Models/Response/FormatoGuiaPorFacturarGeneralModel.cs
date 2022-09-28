using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public class FormatoGuiaPorFacturarGeneralModel
    {
        public bool comentariosEntrega { get; set; }
        public string serienumero { get; set; }
        public string guianumero { get; set; }
        public string Cliente { get; set; }
        public string FacturaNumero { get; set; }
        public DateTime? FacturaFecha { get; set; }
        public string EstadoFacturacion { get; set; }
        public DateTime? FechaDocumento { get; set; }
        public string ReferenciaNumeroOrden { get; set; }
        public string Descripcion { get; set; }
        public decimal Cantidad { get; set; }
        public string LicitacionNumeroProceso { get; set; }
        public string ReprogramacionPuntoPartida { get; set; }
        public string estado { get; set; }
        public DateTime? FechaRecepcion { get; set; }
        public DateTime? Retorno { get; set; }

    }
}
