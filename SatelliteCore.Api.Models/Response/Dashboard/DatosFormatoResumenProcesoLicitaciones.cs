using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.Dashboard
{
    public struct DatosFormatoResumenProcesoLicitaciones
    {
        public int IdProceso { get; set; }
        public string DescripcionProceso { get; set; }
        public int IdDetalle { get; set; }
        public int NumeroItem { get; set; }
        public string DescripcionItem { get; set; }
        public string CodItem { get; set; }
        public decimal PrecioUnitario { get; set; }
        public int Cliente { get; set; }
        public string DcliNombreCompleto { get; set; }
        public int NumeroEntrega { get; set; }
        public int CantidadEntregar { get; set; }
        public decimal MontoCobrar { get; set; }
        public DateTime FechaVencimiento { get; set; }
        public decimal CantidadEntregadaPlazo { get; set; }
        public decimal MontoFacturado { get; set; }
    }
}
