using System;

namespace SatelliteCore.Api.Models.Request
{
    public struct ControlAgujasModel
    {
        public decimal Cantidad { get; set; }
        public DateTime FechaVencimiento { get; set; }
        public int Secuencia { get; set; }
        public string ControlNumero { get; set; }
    }
}
