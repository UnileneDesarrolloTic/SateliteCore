using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.OCDrogueria
{
    public struct FormatoDatosCabeceraOrdenCompraPrevio
    {
        public int Proveedor { get; set; }
        public string Clasificacion { get; set; }
        public string DescripcionProveedor { get; set; }
        public string Procedencia { get; set; }
        public string MonedaCodigo { get; set; }
        public DateTime FechaPreparacion { get; set; }
        public decimal MontoTotal { get; set; }
        public string Estado { get; set; }
        public int IdGestionarColor { get; set; }
        public int DiasEspera { get; set; }
        public DateTime FechaPrometida { get; set; }
    }
}
