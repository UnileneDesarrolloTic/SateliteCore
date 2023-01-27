using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.Contabilidad
{
    public struct FormatoDatosCierreHistorico
    {
        public int Id { get; set; }
        public string Periodo { get; set; }
        public decimal CantidadTotalAntes { get; set; }
        public decimal MontoTotalAntes { get; set; }
        public string Tipo { get; set; }
        public decimal MontoTotalActual { get; set; }
        public decimal CantidadTotalActual { get; set; }
        public bool Comparacion { get; set; }
        public int CantidadDiferencia { get; set; }
        public int MontoDiferencia { get; set; }
    }
}
