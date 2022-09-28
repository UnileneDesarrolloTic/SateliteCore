using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request
{
    public struct DatosFormatoKardexInternoGCM
    {
       public int IdKardex { get; set; }
       public string NumeroLote { get; set; }
       public string OrdenFabricacion { get; set; }
       public string TipoTransaccion { get; set; }
       public decimal Cantidad { get; set; }
       public string Usuario { get; set; }
       public string Comentarios { get; set; }
       public string Estado { get; set; }
    }
}
