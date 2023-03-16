using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.Logistica;
using SatelliteCore.Api.ReportServices.Contracts.Logistica;
using SatelliteCore.Api.Services.Contracts;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services
{
    public class LogisticaServices : ILogisticaServices
    {
        private readonly ILogisticaRepository _logisticaRepository;

        public LogisticaServices(ILogisticaRepository logisticaRepository)
        {
            _logisticaRepository = logisticaRepository;
        }
        public async Task<IEnumerable<DatosFormatoPlanOrdenServicosD>> ObtenerNumeroGuias(string numeroguia)
        {
            return await _logisticaRepository.ObtenerNumeroGuias(numeroguia);
        }

        public async Task<ResponseModel<string>> ExportarExcelRetornoGuia()
        {
            IEnumerable<DatosFormatosReporteRetornoGuia> result = new List<DatosFormatosReporteRetornoGuia>();
            result = await _logisticaRepository.ExportarExcelRetornoGuia();
            if(result.Count()==0)
               return new ResponseModel<string>(false, "No hay datos para exportar", "");

            ReporteRetornoGuias_Excel ExporteReporteRetornoGuiaExcel = new ReporteRetornoGuias_Excel();
            string reporte = ExporteReporteRetornoGuiaExcel.GenerarReporte(result);
            return new ResponseModel<string> (true, Constante.MESSAGE_SUCCESS, reporte);
        }

        public async Task<ResponseModel<string>> RegistrarRetornoGuia(List<DatosFormatoRetornoGuiaRequest> dato)
        {
            await _logisticaRepository.RegistrarRetornoGuia(dato);

            return new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Se guardo los datos");
        }

        public async Task<IEnumerable<DatosFormatoItemVentas>> ListarItemVentas(FormatoDatosBusquedaItemsVentas dato)
        {
            return await _logisticaRepository.ListarItemVentas(dato);
        }

        public async Task<IEnumerable<DatosFormatoItemLoteAlmacen>> BuscarItemVentas(string Item)
        {
            return await _logisticaRepository.BuscarItemVentas(Item);
        }

        public async Task<ResponseModel<string>> ListarItemVentasExportar(FormatoDatosBusquedaItemsVentas dato)
        {
            IEnumerable<DatosFormatoItemVentas> result = new List<DatosFormatoItemVentas>();
            result = await _logisticaRepository.ListarItemVentas(dato);

            ReporteItemVentas ExporteItemventas = new ReporteItemVentas();
            string reporte = ExporteItemventas.GenerarReporte(result);

            ResponseModel<string> Respuesta = new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, reporte);

            return Respuesta;
        }

        public async Task<IEnumerable<DatosFormatoDetalledelItemVentas>> ListarItemVentasDetalle()
        {
            return await _logisticaRepository.ListarItemVentasDetalle();
        }

        public async Task<ResponseModel<string>> ListarItemVentasDetalleExportar()
        {
            IEnumerable<DatosFormatoDetalledelItemVentas> result = new List<DatosFormatoDetalledelItemVentas>();
            result = await _logisticaRepository.ListarItemVentasDetalle();

            ReportItemVentasDetalle ExporteItemventasDetalle = new ReportItemVentasDetalle();
            string reporte = ExporteItemventasDetalle.GenerarReporteDetalle(result);

            ResponseModel<string> Respuesta = new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, reporte);

            return Respuesta;
        }

        public async Task<IEnumerable<DatosFormatoDetalleComprometidoItem>> DetalleComprometidoItem(DatosFormatoRequestDetalleComprometido dato)
        {
            return await _logisticaRepository.DetalleComprometidoItem(dato);
        }

        public async Task<IEnumerable<DatosFormatoMateriaPrimaItemLogistica>> BuscarNumeroPedido(string NumeroDocumento, string Tipo)
        {
            return await _logisticaRepository.BuscarNumeroPedido(NumeroDocumento, Tipo);
        }

        public async Task<IEnumerable<DatosFormatoDetalleRecetaMPLogistica>> BuscardDetalleRecetaMP(string Item, string Cantidad)
        {
            return await _logisticaRepository.BuscardDetalleRecetaMP(Item, Cantidad);
        }

    }
}
