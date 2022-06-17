using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public class DatosFormatoDistribuccionLP
    {
        public int idEntrega { get; set; }
        public string DescripcionProceso { get; set; }
        public string NombreDiresa { get; set; }
        public string CodigoAlmacen { get; set; }
        public string PuntosdeEntrega { get; set; }
        public string Tipodeusuario { get; set; }
        public int NumeroItem { get; set; }
        public string CodItem { get; set; }
        public int CodigoSISMED { get; set; }
        public string DescripcionItem { get; set; }
        public decimal PrecioUnitario { get; set; }
        public decimal CantidadRequerida { get; set; }
        public string OrdenCompra { get; set; }
        public int Cantidad { get; set; }
        public string Pecosa { get; set; }

    }
}
