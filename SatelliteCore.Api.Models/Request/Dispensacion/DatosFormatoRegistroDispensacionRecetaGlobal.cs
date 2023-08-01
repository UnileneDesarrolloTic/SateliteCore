using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request.Dispensacion
{
    public struct DatosFormatoRegistroDispensacionRecetaGlobal
    {
         public decimal cantidadDespachada { get; set; }
         public decimal cantidadGeneral { get; set; }
         public decimal cantidadSolicitada { get; set; }
         public decimal cantidadIngresada { get; set; }
         public string descripcionLocal { get; set; }
         public string itemTerminado { get; set; }
         public string documento { get; set; }
         public int entregadoPor { get; set; }
         public string itemInsumo { get; set; }
         public string itemTipo { get; set; }
         public string lote { get; set; }
         public string numeroLote { get; set; }
         public string ordenFabricacion { get; set; }
         public string recibidoPor { get; set; }
         public int secuencia { get; set; }
         public string tipoMP { get; set; }
         public string unidadCodigo { get; set; }
    }
}
