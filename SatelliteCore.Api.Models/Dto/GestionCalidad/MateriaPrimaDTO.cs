using System;

namespace SatelliteCore.Api.Models.Dto.GestionCalidad
{
    public struct MateriaPrimaDTO
    {
        public string TipoDocumento { get; set; }
        public string NumeroDocumento { get; set; }
        public DateTime FechaDocumento { get; set; }
        public string Transaccion { get; set; }
        public string Item { get; set; }
        public string Descripcion { get; set; }
        public string AlmacenCodigo { get; set; }
        public string ReferenciaTipoDocumento { get; set; }
        public string ReferenciaNumeroDocumento { get; set; }
        public string OrdenFabricacion { get; set; }
        public string Estado { get; set; }
        public string NumeroAnalisis { get; set; }
        public decimal Cantidad { get; set; }
    }
}
