using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.OCDrogueria
{
    public class DatosFormatoMostrarOrdenCompraVencidas
    {
        public string Item { get; set; }
        public string NumeroOrden { get; set; }
        public decimal CantidadPedida { get; set; }
        public decimal CantidadRecibida { get; set; }
        public string Comentario { get; set; }
        public string Excluir { get; set; }
    }
}
