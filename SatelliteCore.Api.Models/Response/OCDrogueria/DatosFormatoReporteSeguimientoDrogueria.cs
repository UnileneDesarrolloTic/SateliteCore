using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.OCDrogueria
{
    public struct DatosFormatoReporteSeguimientoDrogueria
    {
        public int IdProveedor { get; set; }
        public string DescProveedor { get; set; }
        public string Puerto { get; set; }
        public string Periodo { get; set; }
        public string Item { get; set; }
        public decimal VolumenCaja { get; set; }
        public decimal Variacion { get; set; }
        public decimal Pronostico { get; set; }
        public decimal DiasAdicionales { get; set; }
        public decimal ConsumoDiario { get; set; }
        public int VariableItem { get; set; }
        public int TGestionCompra { get; set; }
        public int TGestionPago { get; set; }
        public int TGestionAprobacion { get; set; }
        public int Transporte { get; set; }
        public int TAduanas { get; set; }
        public int MesesAdicional { get; set; }
        public decimal StockAlmacenDRO { get; set; }
        public decimal CantidadOC { get; set; }
        public decimal totalStock { get; set; }
        public int FuturoStock { get; set; }
        public int PuntoCorteDebePagar { get; set; }
        public int MaximoStock { get; set; }
    }
}
