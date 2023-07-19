using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.Dispensacion
{
    public struct DatosFormatoObtenerOrdenFabricacion
    {
        public string OrdenFabricacion { get; set; }
        public string Lote { get; set; }
        public DateTime FechaProduccion { get; set; }
        public DateTime FechaExpiracion { get; set; }
        public string Item { get; set; }
        public string Numerodeparte { get; set; }
        public string DescripcionLocal { get; set; }
        public string Busqueda { get; set; }
        public decimal CantidadProgramada { get; set; }
        public decimal CantidadPedida { get; set; }
        public decimal CantidadMuestra { get; set; }
        public decimal CantidadProducida { get; set; }
        public string Estado { get; set; }
        public string ReferenciaTipo { get; set; }
        public int Cliente { get; set; }    
        public string TransferidoFlag { get; set; }
        public string Comentarios { get; set; }
        public DateTime? FechaEntrega { get; set; }
        public DateTime? FechaRequerida { get; set; }
        public string PedidoNumero { get; set; }
        public int Multipedido { get; set; }
        public DateTime? FechaProgramadaInicio { get; set; }
        public int Secuencia { get; set; }
    }
}
