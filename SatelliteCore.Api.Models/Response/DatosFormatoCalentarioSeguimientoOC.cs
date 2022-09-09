using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatoCalentarioSeguimientoOC
    {
        public string Item { get; set; }
        public string DescripcionLocal { get; set; }
        public string UnidadCodigo { get; set; }
        public int Pronostico { get; set; }
        public decimal CoeficienteVariacion { get; set; }
        public int LimiteSuperior { get; set; }
        public int PuntoControl { get; set; }
        public int Maximo { get; set; }
        public int Aduanas { get; set; }
        public decimal PendienteOC { get; set; }
        public decimal TransporteCC { get; set; }
        public decimal StockActual { get; set; }
        public decimal StockComprometido { get; set; }
        public int StockDisponible { get; set; }
        public decimal CantPedir { get; set; }
        public string January { get; set; }
        public string February { get; set; }
        public string March { get; set; }
        public string April { get; set; }
        public string May { get; set; }
        public string June { get; set; }
        public string July { get; set; }
        public string August { get; set; }
        public string September { get; set; }
        public string October { get; set; }
        public string November { get; set; }
        public string December { get; set; }
    }
}
