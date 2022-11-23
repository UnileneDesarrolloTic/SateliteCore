using System;
using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Request.GestionCalidad
{
    public struct CabeceraDetalleReclamoDTO
    {
        public int Cliente { get; set; }
        public DateTime FechaRegistro { get; set; }
        public string RazonSocial { get; set; }
        public string Documento { get; set; }
        public string Pais { get; set; }
        public string Territorio { get; set; }
        public string Estado { get; set; }
    }

    public struct DetalleReclamoDTO
    {
        public string Documento { get; set; }
        public string Lote { get; set; }
        public string OrdenFabricacion { get; set; }
        public string Linea { get; set; }
        public string Familia { get; set; }
        public string Item { get; set; }
        public string DescripcionItem { get; set; }
        public string Marca { get; set; }
        public decimal Cantidad { get; set; }
        public string Estado { get; set; }
        public string UsuarioRegistro { get; set; }
        public DateTime FechaRegistro { get; set; }

    }
    public struct ReclamoDTO
    {
        public CabeceraDetalleReclamoDTO Cabecera { get; set; }
        public List<DetalleReclamoDTO> Detalle { get; set; }
    }
}
