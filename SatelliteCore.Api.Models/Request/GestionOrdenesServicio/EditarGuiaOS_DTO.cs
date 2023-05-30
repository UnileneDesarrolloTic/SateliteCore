using System;

namespace SatelliteCore.Api.Models.Request.GestionOrdenesServicio
{
    public struct EditarGuiaOS_DTO
    {
        public int Id { get; set; }
        public int Cabecera { get; set; }
        public DateTime Fecha { get; set; }
        public string Cliente { get; set; }
        public string Direccion { get; set; }
        public string Departamento { get; set; }
        public string Comercial { get; set; }
        public decimal Peso { get; set; }
        public int Bultos { get; set; }
        public string Comentario { get; set; }
        public string Usuario { get; set; }

        public bool Validar()
        {
            if (Cabecera < 1 || string.IsNullOrWhiteSpace(Cliente) || string.IsNullOrWhiteSpace(Direccion) || Id == 0)
                return false;

            return true;
        }
    }
}
