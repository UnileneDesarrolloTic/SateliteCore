using SatelliteCore.Api.Models.Dto.AnalisisAgujas;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Request.GestionGuias;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.Logistica;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services.Contracts
{
    public interface ILogisticaServices
    {
        public Task<ResponseModel<DatosFormatoPlanOrdenServicosD>> ObtenerNumeroGuias(string numeroguia);
        public Task<ResponseModel<string>> RegistrarRetornoGuia(DatosFormatoRetornoGuiaRequest dato);
        public Task<ResponseModel<IEnumerable<DatosFormatosReporteRetornoGuia>>> ListarRetornoGuia(DatosFormatoGestionGuiasClienteModel dato);
        public Task<ResponseModel<string>> ExportarExcelRetornoGuia(DatosFormatoGestionGuiasClienteModel dato);
        public Task<IEnumerable<DatosFormatoItemVentas>> ListarItemVentas(FormatoDatosBusquedaItemsVentas dato);
        public Task<IEnumerable<DatosFormatoItemLoteAlmacen>> BuscarItemVentas(string Item);
        public Task<ResponseModel<string>> ListarItemVentasExportar(FormatoDatosBusquedaItemsVentas dato);
        public Task<IEnumerable<DatosFormatoDetalledelItemVentas>> ListarItemVentasDetalle();
        public Task<ResponseModel<string>> ListarItemVentasDetalleExportar();
        public Task<IEnumerable<DatosFormatoDetalleComprometidoItem>> DetalleComprometidoItem(DatosFormatoRequestDetalleComprometido dato);
        public Task<IEnumerable<DatosFormatoMateriaPrimaItemLogistica>> BuscarNumeroPedido(string NumeroDocumento, string Tipo);
        public Task<IEnumerable<DatosFormatoDetalleRecetaMPLogistica>> BuscardDetalleRecetaMP(string Item, string Cantidad);

    }
}
