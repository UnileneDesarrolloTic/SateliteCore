using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request
{
    
    public class FormatoComprobantePagoDetraccion
    {   
        public string TipoCuenta { get; set; }
        public string NumeroCuenta { get; set; }
        public string NumeroConstancia { get; set; }
        public string PeriodoTributario { get; set; }
        public string RucProveedor { get; set; }
        public string NombreProveedor { get; set; }
        public string TipoDocumento { get; set; }
        public string DocumentoAdquiriente { get; set; }
        public string RazonSocial { get; set; }
        public DateTime FechaPago { get; set; }
        public decimal MontoDeposito { get; set; }
        public string TipoBien { get; set; }
        public string TipoOperacion { get; set; }
        public string TipodeComprobante { get; set; }
        public string Serie { get; set; }
        public string Numero { get; set; }
        public string PagoDetraccion { get; set; }
    }
}
