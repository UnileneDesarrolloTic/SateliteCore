using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DCompraMPArimaModel
    {
        public string Item { get; set; }
        public int CantidadPedida { get; set; }
        public int DetallePendienteOC { get; set; }
        public string NumeroOrden { get; set; }
        public DateTime FechaPrometida { get; set; }
        public string EstadoDetalle { get; set; }
        public int idProveedor { get; set; }
        public string NombreProveedor { get; set; }

    }
}
