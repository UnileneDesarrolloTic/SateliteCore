using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.Dispensacion
{
    public struct  DatosFormatoDispensacionRecetaDetalle
    {
        public string OrdenFabricacion { get; set; }
        public string NumeroLote { get; set; }
        public string ItemTerminado { get; set;  }
        public int Secuencia { get; set; }
        public string Documento { get; set; }
        public string ItemInsumo { get; set; }
        public string DescripcionLocal { get; set; }
        public string ItemTipo { get; set; }
        public string UnidadCodigo { get; set; }
        public decimal CantidadGeneral { get; set; }
        public decimal CantidadSolicitada { get; set; }
        public decimal CantidadDespachada { get; set; }
        public string TipoMP { get; set; }
        public string Lote { get; set; }
        public int EntregadoPor { get; set; }
        public string RecibidoPor { get; set; }
        public decimal CantidadIngresada { get; set; }
    }
}
