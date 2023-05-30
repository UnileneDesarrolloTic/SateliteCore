using System;

namespace SatelliteCore.Api.Models.Request.GestionOrdenesServicio
{
    public struct ListarOrdenServicioResponseDTO
    {
        public int Id { get; set; }
        public DateTime Fecha { get; set; }
        public DateTime? FechaSalida { get; set; }
        public string OrdenServicio { get; set; }
        public int IdTransportista { get; set; }
        public string Transportista { get; set; }
        public string Usuario { get; set; }
        public string Estado { get; set; }
    }
}
