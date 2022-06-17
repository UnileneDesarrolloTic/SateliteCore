using System;

namespace SatelliteCore.Api.Models.Dto.AnalisisAgujas
{
    public struct ObtenerDatosGeneralesDTO
    {
        public string CodTipo { get; set; }
        public string Tipo { get; set; }
        public string CodLongitud { get; set; }
        public string Longitud { get; set; }
        public string CodBroca { get; set; }
        public string Broca { get; set; }
        public string CodAlambre { get; set; }
        public string Alambre { get; set; }
        public string Serie { get; set; }
        public string OrdenCompra { get; set; }
        public string ControlNumero { get; set; }
        public string Proveedor { get; set; }
        public int Cantidad { get; set; }
        public int UndMuestrear { get; set; }
        public int UndMuestrearI { get; set; }
        public int UndMuestrearIII { get; set; }
        public decimal FuerzaPerforacion { get; set; }
        public string Observaciones { get; set; }
        public DateTime FechaAnalisis { get; set; }
    }
}
