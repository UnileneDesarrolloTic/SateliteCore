using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.Dispensacion
{
    public struct DatosFormatoHistorialDispensaccion
    {
        public string Documento { get; set; }
        public int Secuencia { get; set; }
        public string Item { get; set; }
        public string Lote { get; set; }
        public decimal CantidadIngresada { get; set; }
        public decimal CantidadSolicitada { get; set; }
        public string UsuarioCreacion { get; set; }
        public DateTime FechaRegistro { get; set; }
    }
}
