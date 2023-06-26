using System;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatoPlanOrdenServicosD
    {
        public string NumeroGuia { get; set; }
        public DateTime FechaDocumento { get; set; }
        public string Cliente { get; set; }
        public string OrdenServicios { get; set; }
        public DateTime FechaRetorno { get; set; }
        public DateTime FechaOrdenServicio { get; set; }
        public string OrdenServicio { get; set; }
        public string NombreTransportista { get; set; }
    }
}
