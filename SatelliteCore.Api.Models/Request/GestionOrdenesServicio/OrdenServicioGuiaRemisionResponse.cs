using System;

namespace SatelliteCore.Api.Models.Request.GestionOrdenesServicio
{
    public struct OrdenServicioGuiaRemisionResponse
    {
        public string Guia { get; set; }
        public DateTime FechaGuia { get; set; }
        public string DocumentoCliente { get; set; }
        public string Cliente { get; set; }
        public string Departamento { get; set; }
        public string DestinatarioDireccion { get; set; }
        public string Documento { get; set; }
        public DateTime? FechaDocumento { get; set; }
        public string Almacen { get; set; }
        public string EstadoGuia { get; set; }
        public string EstadoFacturacion { get; set; }
        public string Usuario { get; set; }

    }
}
