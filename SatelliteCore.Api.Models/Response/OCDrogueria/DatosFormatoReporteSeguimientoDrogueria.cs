using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.OCDrogueria
{
    public struct DatosFormatoReporteSeguimientoDrogueria
    {
        public int IdProveedor { get; set; }
        public int OrdenAgrupador { get; set; }
        public string DescProveedor { get; set; }
        public string Puerto { get; set; }
        public string Periodo { get; set; }
        public string Item { get; set; }
        public string DescripcionLocal { get; set; }
        public string UnidadCodigo { get; set; }
        public decimal VolumenCaja { get; set; }
        public decimal Variacion { get; set; }
        public decimal Pronostico { get; set; }
        public decimal DiasAdicionales { get; set; }
        public int TiempoTotal { get; set; }
        public int TiempoGeneral { get; set; }
        public decimal ConsumoDiario { get; set; }
        public int TFabricacion { get; set; }
        public int TGestionCompra { get; set; }
        public int TGestionPago { get; set; }
        public int TGestionAprobacion { get; set; }
        public int Transporte { get; set; } 
        public int TAduanas { get; set; }
        public int MesesAdicional { get; set; }
        public int StockAlmacenDRO { get; set; }
        public int CantidadOC { get; set; }
        public int TotalStock { get; set; }
        public int FuturoStock { get; set; }
        public int DiasEspera { get; set; }
        public int PuntoCorteDebePagar { get; set; }
        public int MaximoStock { get; set; }
        public int CantidadComprar { get; set; }
        public int TotalComprar { get; set; }
        public decimal PrecioFOBFinal { get; set; }
        public decimal PrecioFOBTotalFinal { get; set; }
        public string ColorVariacion { get; set; }
        public int IdGestionarColor { get; set; }

    }
}
