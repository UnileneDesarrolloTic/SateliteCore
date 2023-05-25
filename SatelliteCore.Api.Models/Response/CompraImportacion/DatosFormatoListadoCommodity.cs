using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.CompraImportacion
{
    public class DatosFormatoListadoCommodity
    {
        public string Material { get; set; }
        public string ItemFinal { get; set; }
        public string DescripcionLocal { get; set; }
        public string Clasificacion { get; set; }
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
        public int PendienteOC { get; set; }
        public int PreparacionOC { get; set; }
        public int Almacen { get; set; }
        public int Disponible { get; set; }
        public int DiasPotencial { get; set; }
        public int PuntoCorte { get; set; }
        public int MaximoStock { get; set; }
        public int DiasEspera { get; set; }
        public decimal CantidadComprar { get; set; }
        public string GestionColor { get; set; }
        public int IdGestion { get; set; }
    }
}
