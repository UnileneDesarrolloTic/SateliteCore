using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatosListarControlLotes
    {
        public DateTime FechaEntrega { get; set; }
        public string Aprobado { get; set; }
        public string CompaniaSocio { get; set; }
        public string OrdenFabricacion { get; set; }
        public string Lote { get; set; }
        public DateTime FechaRegistro { get; set; }
        public DateTime FechaProduccion { get; set; }
        public string PeriodoProduccion { get; set; }
        public string Item { get; set; }
        public decimal CantidadProgramada { get; set; }
        public decimal CantidadPedida { get; set; }
        public decimal CantidadProducida { get; set; }
        public string Estado { get; set; }
        public string UltimoUsuario  { get; set; }
        public DateTime UltimaFechaModif  { get; set; }
        public decimal DiaProduccion { get; set; }
        public string DescripcionLocal  { get; set; }
        public string ItemTipo { get; set; }
        public string UnidadCodigo { get; set; }
        public string SeccionPlanta  { get; set; }
        public string ReferenciaTipo { get; set; }
        public string ReferenciaNumero  { get; set; }
        public string AuditableFlag  { get; set; }
        public decimal CantidadMuestra  { get; set; }
        public int Cliente { get; set; }
        public string TransferidoFlag { get; set; }
        public string Comentarios  { get; set; }
        public DateTime FechaRequerida  { get; set; }
        public string PedidoNumero  { get; set; }
        public string NumerodeParte  { get; set; }
        public string Busqueda  { get; set; }
        public DateTime FechaExpiracion  { get; set; }
        public int ProtocoloFlag { get; set; }
    }
}
