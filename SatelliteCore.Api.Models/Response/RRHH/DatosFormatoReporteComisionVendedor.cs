using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.RRHH
{
    public class DatosFormatoReporteComisionVendedor
    {
        public int Cobrador { get; set; }
        public string NombreCompleto { get; set; }
        public decimal TotalFacturado { get; set; }
        public decimal SinIGV { get; set; }
        public decimal VentaGeneral { get; set; }
        public decimal VentaGuantes { get; set; }
        public decimal VentaEmodialisis { get; set; }
        public decimal ComisionGeneral { get; set; }
        public decimal ComisionGuantes { get; set; }
        public decimal ComisionEmodialisis { get; set; }
        public decimal ComisionTotal { get; set; }
    }
}
