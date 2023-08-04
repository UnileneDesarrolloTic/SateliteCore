using System;

namespace SatelliteCore.Api.Models.Encajado
{
    public class ListaOrdenesFabricaciónDTO
    {
        public string OrdenFabricacion { get; set; }
        public string Item { get; set; }
        public string Descripcion { get; set; }
        public decimal CantidaProgramada { get; set; }
        public decimal CantidadProducida { get; set; }
        public decimal CantidadMuestra { get; set; }
        public string Estado { get; set; }
        public DateTime? FechaExpiracion { get; set; }
    }
}
