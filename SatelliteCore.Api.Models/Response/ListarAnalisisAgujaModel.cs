using System;

namespace SatelliteCore.Api.Models.Response
{
    public struct ListarAnalisisAgujaModel
    {
        public string Lote { get; set; }
        public string Item { get; set; }
        public string DescripcionItem { get; set; }
        public string OrdenCompra { get; set; }
        public string Proveedor { get; set; }
        public int Cantidad { get; set; }
        public DateTime FechaAnalisis { get; set; }
        public int CantidadPruebas { get; set; }
    }
}
