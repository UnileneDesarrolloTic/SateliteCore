using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Request.GestionOrdenesServicio
{
    public struct DatosRegistrarOrdenServicioDTO
    {
        public List<OrdenServicioDetalle> Detalle { get; set; }
        public int Transportista { get; set; }
        public string Usuario { get; set; }
    }
}
