using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Response
{
    public class CReporteGuiaRemisionModel
    {
        public string DescripcionProceso { get; set; }
        public string DescripcionComercialDetalle { get; set; }
        public int CantItems { get; set; }
        public string Region { get; set; }
        public string OrdenCompra { get; set; }
        public string Pecosa { get; set; }
        public string Contrato { get; set; }
        public string NumeroEntrega { get; set; }
        public string ClienteNombre { get; set; }
        public string NombreRegion { get; set; }
        public string GuiaNumero { get; set; }

        public List<DReportGuiaRemisionModel> DetalleGuia { get; set; }

        public CReporteGuiaRemisionModel()
        {
            DetalleGuia = new List<DReportGuiaRemisionModel>();
        }
    }
}
