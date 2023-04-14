using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.CompraAguja
{
    public struct DatosFormatoListadoSeguimientoCompraAguja
    {
        public string Material { get; set; }
        public string ItemFinal { get; set; }
        public string DescripcionLocal { get; set; }
        public string LongAguja { get; set; }
        public string FamiliaLarga { get; set; }
        public decimal TiempoCompra { get; set; }
        public decimal TiempoPago { get; set; }
        public decimal TiempoAprobacion { get; set; }
        public decimal Tiempofabricacion { get; set; }
        public decimal TiempoTransporte { get; set; }
        public decimal TiempoAduanas { get; set; }
        public decimal CantidadMinima { get; set; }
        public int Pronostico { get; set; }
        public decimal Variacion { get; set; }
        public decimal DemoraLlegarProducto { get; set; }
        public decimal ConsumoDia { get; set; }
        public decimal DesviacionCompra { get; set; }
        public decimal PendienteOC { get; set; }
        public decimal ControlCalidad { get; set; }
        public decimal Aduanas { get; set; }
        public decimal Almacen { get; set; }
        public decimal Disponible { get; set; }
        public int DiasPotencial { get; set; }
        public int PuntoCorte { get; set; }
        public int MaximoStock { get; set; }
        public int DiasEspera { get; set; }
        public decimal CantidadComprar { get; set; }

    }
}
