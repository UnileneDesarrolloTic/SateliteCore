using System;

namespace SatelliteCore.Api.Models.Entities
{
    public struct TBDReclamosEntity
    {
        public int IdDetalle { get; set; }
        public string CodReclamo { get; set; }
        public string Lote { get; set; }
        public string OrdenFabricacion { get; set; }
        public string TipoDocumento { get; set; }
        public string  Item { get; set; }
        public string Documento { get; set; }
        public string Motivo { get; set; }
        public string Solicitud { get; set; }
        public string Remitente { get; set; }
        public bool Reingreso { get; set; }
        public bool PorCarta { get; set; }
        public DateTime FechaIncidencia { get; set; }
        public int? Clasificacion { get; set; }
        public int? AreaInvolucrada { get; set; }
        public decimal Cantidad { get; set; }
        public string Observaciones { get; set; }
        public string Estado { get; set; }
        public string UsuarioRegistro { get; set; }
        public DateTime FechaRegistro { get; set; }

        public bool ValidarCreacion()
        {
            if (IdDetalle != 0)
                return false;

            if(string.IsNullOrEmpty(CodReclamo) || string.IsNullOrEmpty(Lote) || string.IsNullOrEmpty(OrdenFabricacion) || string.IsNullOrEmpty(TipoDocumento) ||
                string.IsNullOrEmpty(Documento) || string.IsNullOrEmpty(Item) || string.IsNullOrEmpty(Motivo) || string.IsNullOrEmpty(Solicitud) || 
                string.IsNullOrEmpty(UsuarioRegistro) || string.IsNullOrEmpty(Estado) || string.IsNullOrEmpty(Remitente))
                return false;

            return true;
        }
    }
}
