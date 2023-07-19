using SatelliteCore.Api.CrossCutting.Helpers;
using System;

namespace SatelliteCore.Api.Models.Entities
{
    public struct TBMAnalisisHebraEntity
    {
        public string OrdenCompra { get; set; }
        public string NumeroAnalisis { get; set; }
        public DateTime FechaAnalisis { get; set; }
        public string Certificado { get; set; }
        public string Quimica { get; set; }
        public string Conclusion { get; set; }
        public string Observaciones { get; set; }
        public string UsuarioRegistro { get; set; }
        public DateTime? FechaRegistro { get; set; }
        public string UsuarioModificacion { get; set; }
        public DateTime? FechaModificacion { get; set; }

        public bool ValidarDatos()
        {
            if(string.IsNullOrWhiteSpace(OrdenCompra) || string.IsNullOrWhiteSpace(NumeroAnalisis) || string.IsNullOrWhiteSpace(Certificado) || 
                string.IsNullOrWhiteSpace(Quimica) || string.IsNullOrWhiteSpace(Conclusion) 
                || !Shared.ValidarFecha(FechaAnalisis) || !Shared.ValidarFecha(FechaRegistro) )
                return false;

            return true;
        }
    }
}
