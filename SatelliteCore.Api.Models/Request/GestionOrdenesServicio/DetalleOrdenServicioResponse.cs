using System;

namespace SatelliteCore.Api.Models.Request.GestionOrdenesServicio
{
    public struct DetalleOrdenServicioResponse
    {
        public int Id { get; set; }
        public string Guia { get; set; }
        public DateTime Fecha { get; set; }
        public string Cliente { get; set; }
        public string Direccion { get; set; }
        public string Departamento { get; set; }
        public string Factura { get; set; }
        public decimal Peso { get; set; }
        public int Bultos { get; set; }
        public string Comentario { get; set; }
    }
}
