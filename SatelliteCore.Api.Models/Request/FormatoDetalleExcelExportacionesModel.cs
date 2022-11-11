using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request
{
    public class FormatoDetalleExcelExportacionesModel
    {
        public string Codsut { get; set; }
        public int Cantidad { get; set; }
        public decimal Punitario { get; set; }
    }
}
