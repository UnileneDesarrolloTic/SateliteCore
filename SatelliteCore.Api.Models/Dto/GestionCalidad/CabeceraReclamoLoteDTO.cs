using System;

namespace SatelliteCore.Api.Models.Dto.GestionCalidad
{
    public struct CabeceraReclamoLoteDTO
    {
        public int Id { get; set; }
        public string OrdenFabricacion { get; set; }
        public DateTime FechaDocumento { get; set; }
        public string CodTipoDocumento { get; set; }
        public string TipoDocumento { get; set; }
        public string Documento { get; set; }
        public string Item { get; set; }
        public string Descripcion { get; set; }
        public string Linea { get; set; }
        public string Familia { get; set; }
        public string SubFamilia { get; set; }
        public string Marca { get; set; }
        public decimal CantidadPedida { get; set; }
        public decimal CantidadEntregada { get; set; }
        public decimal PrecioUnitario { get; set; }
        public decimal Monto { get; set; }
        public string Motivo { get; set; }
        public string Remitente { get; set; }
        public bool Reingreso { get; set; }
        public bool PorCarta { get; set; }
        public string Solicitud { get; set; }
        public DateTime FechaIncidencia { get; set; }
        public int Clasificacion { get; set; }
        public int AreaInvolucrada { get; set; }
        public decimal Cantidad { get; set; }
        public string Observaciones { get; set; }



        public string Estado { get; set; }
        public string TipoEnvioRespuesta { get; set; }
        public string DestinatarioRespuesta { get; set; }
        public string LoteCanjeRespuesta { get; set; }
        public string Respuesta { get; set; }
        public string UsuarioRespuesta { get; set; }
        public string FechaRespuesta { get; set; }
    }
}
