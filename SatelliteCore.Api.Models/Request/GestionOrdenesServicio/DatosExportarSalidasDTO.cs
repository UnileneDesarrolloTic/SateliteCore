using System;

namespace SatelliteCore.Api.Models.Request.GestionOrdenesServicio
{
    public struct DatosExportarSalidasDTO
    {
        public string Guia { get; set; }
        public string FechaGuia { get; set; }
        public string Cliente { get; set; }
        public string Direccion { get; set; }
        public string Departamento { get; set; }
        public string Comercial { get; set; }
        public decimal Peso { get; set; }
        public int Bultos { get; set; }
        public string Transportista { get; set; }
        public string FechaRetorno { get; set; }
        public string OrdServicio { get; set; }
        public string FechaServicio { get; set; }
    }
}
