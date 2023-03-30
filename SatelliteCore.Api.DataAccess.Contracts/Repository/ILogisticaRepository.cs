using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Request.GestionGuias;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.Logistica;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Contracts.Repository
{
    public interface ILogisticaRepository
    {
        public Task<DatosFormatoPlanOrdenServicosD> ObtenerNumeroGuias(string numeroguia);
        public Task<int> RegistrarRetornoGuia(DatosFormatoRetornoGuiaRequest dato);
        public Task<IEnumerable<DatosFormatosReporteRetornoGuia>> ListarRetornoGuia(DatosFormatoGestionGuiasClienteModel dato);
        public Task<IEnumerable<DatosFormatoItemVentas>> ListarItemVentas(FormatoDatosBusquedaItemsVentas dato);
        public Task<IEnumerable<DatosFormatoItemLoteAlmacen>> BuscarItemVentas(string Item);
        public Task<IEnumerable<DatosFormatoDetalledelItemVentas>> ListarItemVentasDetalle();
        public Task<IEnumerable<DatosFormatoDetalleComprometidoItem>> DetalleComprometidoItem(DatosFormatoRequestDetalleComprometido dato);
        public Task<IEnumerable<DatosFormatoMateriaPrimaItemLogistica>> BuscarNumeroPedido(string NumeroDocumento, string Tipo);
        public Task<IEnumerable<DatosFormatoDetalleRecetaMPLogistica>> BuscardDetalleRecetaMP(string Item, string Cantidad);

    }
}
