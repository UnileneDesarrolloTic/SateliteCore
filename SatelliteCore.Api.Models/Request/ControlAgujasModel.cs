using System;
using System.ComponentModel.DataAnnotations;

namespace SatelliteCore.Api.Models.Request
{
    public struct ControlAgujasModel
    {
        [Required]
        public string ControlNumero { get; set; }

        [Required]
        public int Secuencia { get; set; }

        [Required]
        public int CantidadPruebas { get; set; }

        [Required]
        public DateTime FechaVencimiento { get; set; }        
        public int CodUsuarioSesion { get; set; }
    }
}
