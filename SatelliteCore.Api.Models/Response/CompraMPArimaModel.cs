using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Response
{
    public class CompraMPArimaModel
    {

        public string Item { get; set; }
        public string Regla { get; set; }
        public string Descripcion { get; set; }
        public string codFamilia { get; set; }
        public string FamiliaMP { get; set; }
        public decimal CoeficienteVariacion { get; set; }
        public string Kraljic { get; set; }
        public double Meses { get; set; }
        public int Promedioanual { get; set; }
        public int promedioHist { get; set; }
        public int Pronostico { get; set; }
        public int LimiteSuperior { get; set; }
        public int Maximo { get; set; }
        public int PuntoControl { get; set; }
        public decimal ControlCalidad { get; set; }
        public decimal Aduanas { get; set; }
        public decimal PendienteOC { get; set; }
        public decimal StockDisponible { get; set; }
        public decimal Alerta { get; set; }


        public List<DCompraMPArimaModel> DetalleCompra { get; set; }


        public CompraMPArimaModel()
        {
            DetalleCompra = new List<DCompraMPArimaModel>();
        }


    }
}
