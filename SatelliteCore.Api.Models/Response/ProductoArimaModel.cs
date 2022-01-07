using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Response
{
    public class ProductoArimaModel
    {
        public int Candidato { get; set; }
        public string CodSut { get; set; }
        public string Item { get; set; }
        public string Regla { get; set; }
        public string Descripcion { get; set; }
        public decimal CoeficienteVariacion { get; set; }
        public int Pronostico { get; set; }
        public int LimiteSuperior { get; set; }
        public int PuntoControl { get; set; }
        public int StockMax { get; set; }
        public int StockActual { get; set; }
        public int StockComprometido { get; set; }
        public int StockDisponible { get; set; }
        public int StockEnTransito { get; set; }
        public decimal Alerta { get; set; }
        public int DiferenciaMax { get; set; }
        public List<TransitoProductoArimaModel> PedidosTransito { get; set; }

        public ProductoArimaModel()
        {
            PedidosTransito = new List<TransitoProductoArimaModel>();
        }
    }
}
