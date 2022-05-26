using System;


namespace SatelliteCore.Api.Models.Response
{
    public struct DReportGuiaRemisionModel
    {
       public int NumeroItem { get; set; }
       public string Descripcion { get; set; }
       public string CaractervaluesDescripcion { get; set; }
       public string UnidadCodigo { get; set; }
       public decimal CantidadRequerida { get; set; }
       public int Cantidad { get; set; }
       public int CantidadGRD { get; set; }
       public string Guia { get; set; }
       public string Lote { get; set; }
       public DateTime FechaExpiracion { get; set; }
       public string RegistroSanitario { get; set; }
       public string Protocolo { get; set; }
       public string  NumeroMuestreo { get; set; }
       public string  NumeroEnsayo { get; set; }

    }
}
