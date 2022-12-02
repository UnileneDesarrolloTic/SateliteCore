using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Contracts.Repository
{
    public interface ILogisticaRepository
    {
        public Task<IEnumerable<DatosFormatoPlanOrdenServicosD>> ObtenerNumeroGuias(string NumeroGuia);
        public Task<int> RegistrarRetornoGuia(List<DatosFormatoRetornoGuiaRequest> dato);
        public Task<IEnumerable<DatosFormatoItemVentas>> ListarItemVentas(FormatoDatosBusquedaItemsVentas dato);
        public Task<IEnumerable<DatosFormatoItemLoteAlmacen>> BuscarItemVentas(string Item);
        public Task<IEnumerable<DatosFormatoDetalledelItemVentas>> ListarItemVentasDetalle();
        public Task<IEnumerable<DatosFormatoDetalleComprometidoItem>> DetalleComprometidoItem(DatosFormatoRequestDetalleComprometido dato);
        public Task<IEnumerable<DatosFormatoMateriaPrimaItemLogistica>> BuscarNumeroPedido(string NumeroDocumento, string Tipo);
        public Task<IEnumerable<DatosFormatoDetalleRecetaMPLogistica>> BuscardDetalleRecetaMP(string Item, string Cantidad);

    }
}
