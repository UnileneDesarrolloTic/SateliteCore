using System;

namespace SatelliteCore.Api.Models.Encajado
{
    public struct DatosRegistrarAsignacionDTO
    {
        public int IdEncaje { get; set; }
        public int Etapa { get; set; }
        public int Empleado { get; set; }
        public decimal Cantidad { get; set; }
        public DateTime Fecha { get; set; }
        public string UsuarioRegistro { get; set; }

        public bool ValidarDatos()
        {
            if(IdEncaje == 0 || Etapa == 0 || Empleado == 0 || Cantidad < 1 || string.IsNullOrWhiteSpace(UsuarioRegistro))
                return false;
            return true;
        }
    }
}
