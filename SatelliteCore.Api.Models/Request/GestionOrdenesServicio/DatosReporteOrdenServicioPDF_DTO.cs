using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Request.GestionOrdenesServicio
{
    public class DatosReporteOrdenServicioPDF_DTO
    {
        public string OrdenServicio { get; set; }
        public string Transportista { get; set; }
        public string Fecha { get; set; }
        public List<DatosReporteOrdenServicioDetallePDF_DTO> Detalle { get; set; }

        public DatosReporteOrdenServicioPDF_DTO()
        {
            Detalle =  new List<DatosReporteOrdenServicioDetallePDF_DTO>();
        }
    }

    public class DatosReporteOrdenServicioDetallePDF_DTO
    {
        public string Guia { get; set; }
        public string Cliente { get; set; }
        public string Destino { get; set; }
        public string Departamento { get; set; }
        public string Factura { get; set; }
        public decimal Peso { get; set; }
        public int Bultos { get; set; }
    }
}
