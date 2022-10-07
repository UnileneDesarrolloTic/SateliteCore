using SatelliteCore.Api.Models.Dto.AnalisisAgujas;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services.Contracts
{
    public interface ILogisticaServices
    {
        public Task<IEnumerable<DatosFormatoPlanOrdenServicosD>> ObtenerNumeroGuias(string NumeroGuia);
        public Task<ResponseModel<string>> RegistrarRetornoGuia(List<DatosFormatoRetornoGuiaRequest> dato);
        public Task<IEnumerable<DatosFormatoItemVentas>> ListarItemVentas(FormatoDatosBusquedaItemsVentas dato);
        public Task<IEnumerable<DatosFormatoItemLoteAlmacen>> BuscarItemVentas(string Item);
        public Task<ResponseModel<string>> ListarItemVentasExportar(FormatoDatosBusquedaItemsVentas dato);
        public Task<IEnumerable<DatosFormatoDetalledelItemVentas>> ListarItemVentasDetalle();
        public Task<ResponseModel<string>> ListarItemVentasDetalleExportar();
        public Task<IEnumerable<DatosFormatoDetalleComprometidoItem>> DetalleComprometidoItem(DatosFormatoRequestDetalleComprometido dato);
        public Task<IEnumerable<DatosFormatoMateriaPrimaItemLogistica>> BuscarNumeroPedido(string NumeroDocumento, string Tipo);

    }
}
