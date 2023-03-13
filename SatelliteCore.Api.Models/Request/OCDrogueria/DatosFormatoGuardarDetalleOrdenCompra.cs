using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request.OCDrogueria
{
    public struct DatosFormatoGuardarDetalleOrdenCompra
    {
        public int proveedor { get; set; }
        public int secuencia { get; set; }
        public string item { get; set; }
        public string descripcion { get; set; }
        public string presentacion { get; set; }
        public decimal cantidadpedida { get; set; }
        public decimal preciounitario { get; set; }
        public decimal montototal { get; set; }
        public string moneda { get; set; }
        public string estado { get; set; }
        public DateTime fechaPrometida { get; set; }
        public string colorVariacion { get; set; }
        public int idGestionarColor { get; set; }
    }
}
