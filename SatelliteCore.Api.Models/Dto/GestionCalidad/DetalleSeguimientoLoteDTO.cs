using System;
using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Dto.GestionCalidad
{
    public struct OrdenCompraPorLoteDTO
    {
        public string NumeroOrden { get; set; }
        public string ControlNumero { get; set; }
        public string Proveedor { get; set; }
        public string LoteAprobado { get; set; }
        public string Item { get; set; }
        public string UnidadCodigo { get; set; }
        public string Descripcion { get; set; }
        public string NumeroDeParte { get; set; }
        public string ReferenciaTipo { get; set; }
        public decimal CantidadRecibida { get; set; }
        public decimal CantidadAceptada { get; set; }
        public decimal CantidadRechazada { get; set; }
        public decimal CantidadTransferida { get; set; }
        public DateTime FechaAprobacion { get; set; }
        public DateTime FechaVencimiento { get; set; }
        public string Comentario { get; set; }
    }

    public struct OrdenFabricacionPorLoteDTO
    {
        public string OrdenFabricacion { get; set; }
        public string Lote { get; set; }
        public string PedidoNumero { get; set; }
        public string Cliente { get; set; }
        public DateTime FechaProduccion { get; set; }
        public DateTime FechaExpiracion { get; set; }
        public string Item { get; set; }
        public string Descripcion { get; set; }
        public string NumeroDeParte { get; set; }
        public string UnidadCodigo { get; set; }
        public decimal CantidadProducida { get; set; }
        public string Estado { get; set; }
        public decimal DiaProduccion { get; set; }
        public string ReferenciaTipo { get; set; }
        public string AuditableFlag { get; set; }
        public string TransferidoFlag { get; set; }
        public string Comentarios { get; set; }
        public DateTime FechaRequerida { get; set; }
        public int MultiPedido { get; set; }
    }

    public struct DocumentosPedidosDTO
    {
        public string NumeroDocumento { get; set; }
        public string TipoDocumento { get; set; }
        public string ReferenciaDocumentoNC { get; set; }
        public string FormaFacturacion { get; set; }
        public string TipoVenta { get; set; }
        public DateTime FechaDocumento { get; set; }
        public string Ruc { get; set; }
        public string ClienteNombre { get; set; }
        public string ClienteDireccion { get; set; }
        public string Lote { get; set; }
        public string OrdenFabricacion { get; set; }
        public string AlmacenCodigo { get; set; }
        public string Item { get; set; }
        public string Descripcion { get; set; }
        public string Unidad { get; set; }
        public decimal CantidadEntregada { get; set; }
        public decimal PrecioUnitario { get; set; }
        public decimal Monto { get; set; }
        public decimal MontoTotal { get; set; }
    }

    public struct GuiaDocumentos
    {
        public string GuiaNumero { get; set; }
        public string FacturaNumero { get; set; }
        public DateTime FechaDocumento { get; set; }
        public string ReferenciaNumeroPedido { get; set; }
        public string DestinatarioRuc { get; set; }
        public string DestinatarioNombre { get; set; }
        public string DestinatarioDireccion { get; set; }
        public string Lote { get; set; }
        public string ItemCodigo { get; set; }
        public string Descripcion { get; set; }
        public decimal Cantidad { get; set; }
        public decimal CantidadRecibida { get; set; }
    }

    public struct DetalleSeguimientoLoteDTO
    {
        public List<OrdenCompraPorLoteDTO> OrdenesDeCompra { get; set; }
        public List<OrdenFabricacionPorLoteDTO> OrdenesDeFabricacion { get; set; }
        public List<DocumentosPedidosDTO> DocumentosPedidos { get; set; }
        public List<GuiaDocumentos> GuiasRelacionadas { get; set; }
    }
}
