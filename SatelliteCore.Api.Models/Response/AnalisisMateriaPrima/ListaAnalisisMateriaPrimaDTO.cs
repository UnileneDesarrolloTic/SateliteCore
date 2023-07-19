using System;

namespace SatelliteCore.Api.Models.Response.AnalisisMateriaPrima
{
    public struct ListaAnalisisMateriaPrimaDTO
    {
        public string ControlNumero { get; set; }
        public string NumeroOrden { get; set; }
        public string Analisis { get; set; }
        public string Tipo { get; set; }
        public string Item { get; set; }
        public string Descripcion { get; set; }
        public decimal CantidadAceptada { get; set; }
        public DateTime FechaAprobacion { get; set; }
        public string TipoItem { get; set; }
    }
}
