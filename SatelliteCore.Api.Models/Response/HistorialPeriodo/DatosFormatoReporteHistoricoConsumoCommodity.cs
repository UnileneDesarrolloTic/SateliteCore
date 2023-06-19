using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.CompraImportacion
{
    public struct DatosFormatoReporteHistoricoConsumoCommodity
    {

        public string Commodity { get; set; }
        public string DescripcionLocal { get; set; }
        public decimal Meses1 { get; set; }
        public decimal Meses2 { get; set; }
        public decimal Meses3 { get; set; }
        public decimal Meses4 { get; set; }
        public decimal Meses5 { get; set; }
        public decimal Meses6 { get; set; }
        public decimal Meses7 { get; set; }
        public decimal Meses8 { get; set; }
        public decimal Meses9 { get; set; }
        public decimal Meses10 { get; set; }
        public decimal Meses11 { get; set; }
        public decimal Meses12 { get; set; }
        public decimal Desviacion { get; set; }
        public decimal Promedio { get; set; }
        public decimal Variacion { get; set; }
    }
}
