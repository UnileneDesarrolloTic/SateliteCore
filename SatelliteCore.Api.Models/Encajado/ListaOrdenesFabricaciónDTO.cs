using System;

namespace SatelliteCore.Api.Models.Encajado
{
    public class ListaOrdenesFabricaciónDTO
    {
        public string OrdenFabricacion { get; set; }
        public string Lote { get; set; }
        public string CodSut { get; set; }
        public string Item { get; set; }
        public string Descripcion { get; set; }
        public decimal CantidadProgramada { get; set; }
        public string Estado { get; set; }
        public decimal Ingreso { get; set; }
    }
}
