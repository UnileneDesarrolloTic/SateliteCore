using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public class FormatoDetalleCotizacionExportaciones
    {
        public decimal cantidad { get; set; }
        public string codSut { get; set; }
        public int linea { get; set; }
        public string descripcion { get; set; }
        public Boolean iGVExoneradoFlag { get; set; }
        public string item { get; set; }
        public decimal monto { get; set; }
        public decimal precioUnitario { get; set; }
        public decimal receta { get; set; }
        public decimal materiaprima { get; set; }
        public decimal cantidadpedida { get; set; }
    }
}
