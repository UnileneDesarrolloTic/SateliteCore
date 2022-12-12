using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Entities
{
    public class ParametroMastEntity
    {
        public string CompaniaCodigo { get; set; }
        public string AplicacionCodigo { get; set; }
        public string ParametroClave{ get; set; }
        public string DescripcionParametro { get; set; }
        public string Explicacion { get; set; }
        public string TipodeDatoFlag { get; set; }
        public string Texto { get; set; }
        public decimal Numero { get; set; }
        public string Fecha { get; set; }
        public string FinanceComunFlag { get; set; }
        public string Estado { get; set; }
        public string UltimoUsuario { get; set; }
        public DateTime UltimaFechaModif { get; set; }
        public string ExplicacionAdicional { get; set; }
        public string Texto1 { get; set; }
    }
}
