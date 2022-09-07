using System;


namespace SatelliteCore.Api.Models.Response
{
    public struct FormatoEstructuraObtenerOrdenFabricacion
    {
        public DateTime FechaProduccion { get; set; }
        public string Item { get; set; }
        public string NumeroParte { get; set; }
        public string DescripcionLocal { get; set; }
        public string Cliente { get; set; }
        public string OrdenFabricacion { get; set; }
        public string Lote { get; set; }
        public decimal ContraMuestra { get; set; }
        public string NumeroCaja { get; set; }
        public string AuditableFlag { get; set; }
        public string Marca { get; set; }

    }
}
