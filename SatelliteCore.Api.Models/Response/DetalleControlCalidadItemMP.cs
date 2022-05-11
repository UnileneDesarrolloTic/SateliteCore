using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public class DetalleControlCalidadItemMP
    {
        public string Item { get; set; }
        public string Lote { get; set; }
        public string ReferenciaTipoDocumento { get; set; }
        public string ReferenciaNumeroDocumento { get; set; }
        public string TransaccionCodigo { get; set; }
        public decimal Cantidad { get; set; }
        public string ReferenciaTipoDocumentoOrden { get; set; }
        public string ReferenciaNumeroDocumentoOrden { get; set; }
        public string NombreCompleto { get; set; }
    }
}
