using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Report.Cotizacion;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services.Contracts
{
    public interface ICotizacionServices
    {
        public Task<PaginacionModel<CotizacionEntity>> Listar(DatosListarCotizacionesPaginado datos);
        public Task<ObtenerEstructuraFormCotizacionModel> FormatoEstructura(int codFormato);
        public Task<(object cabecera, object detalle)> FormatoDatos(int idFormato, string cotizacion);
        public Task<IEnumerable<ReportesGeneradosPorCotizacionModel>> ReportesPorCotizacion(string cotizacion);
        public Task<ResponseModel<string>> Guardar(ObtenerFormatoCotizacion cotizacion, int usuarioSesion);
        public Task<ResponseModel<string>> ObtenerReporte(string idObject);
        public Task<CotizacionAbstract> ObtenerDatosReporte(string codigoReporte);
        public Task<IEnumerable<FormatosPorClienteModel>> FormatosPorCliente(int idCliente);
        public Task Actualizar(ActualizarReporteCotizacionModel reporte, int usuarioSesion);

        public Task<IEnumerable<ListaFormatoCotizacion>> ListarFormatoCotizacion();
        public Task<IEnumerable<CamposFormatoCotizacionModel>> CamposFormatosCotizacion(int idFormato);


    }
}
