using System;

namespace SatelliteCore.Api.Models.Encajado
{
    public struct AsignacionEncajadoDTO
    {
        public int Id { get; set; }
        public int Empleado { get; set; }
        public string NombreEmpleado { get; set; }
        public DateTime Fecha { get; set; }
        public decimal Cantidad { get; set; }
        public string Estado { get; set; }
    }
}
