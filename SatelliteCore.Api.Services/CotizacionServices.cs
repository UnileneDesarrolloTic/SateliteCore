using MongoDB.Bson;
using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Report.Cotizacion;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.ReportServices.Contracts.Cotizacion;
using SatelliteCore.Api.Services.Contracts;
using System.Collections.Generic;
using System.Threading.Tasks;
using SystemsIntegration.Api.Models.Exceptions;

namespace SatelliteCore.Api.Services
{
    public class CotizacionServices : ICotizacionServices
    {
        private readonly ICotizacionRepository _cotizacionRepository;

        public CotizacionServices(ICotizacionRepository cotizacionRepository)
        {
            _cotizacionRepository = cotizacionRepository;
        }

        public async Task<PaginacionModel<CotizacionEntity>> Listar(DatosListarCotizacionesPaginado datos)
        {
            (List<CotizacionEntity> lista, int totalRegistros) = await _cotizacionRepository.Listar(datos);

            PaginacionModel<CotizacionEntity> response = new PaginacionModel<CotizacionEntity>(lista, datos.Pagina, datos.RegistrosPorPagina, totalRegistros);

            return response;
        }

        public async Task<ObtenerEstructuraFormCotizacionModel> FormatoEstructura(int codFormato)
        {
            ObtenerEstructuraFormCotizacionModel result = await _cotizacionRepository.FormatoEstructura(codFormato);
            return result;
        }

        public async Task<(object cabecera, object detalle)> FormatoDatos(int idFormato, string cotizacion)
        {
            (object cabecera, object detalle) datos = await _cotizacionRepository.FormatoDatos(idFormato, cotizacion);
            return datos;
        }

        public async Task<ResponseModel<string>> Guardar(ObtenerFormatoCotizacion cotizacion, int usuarioSesion)
        {
            BsonDocument documentoBson = BsonDocument.Parse(cotizacion.Cotizacion.ToString());
            string idBson = await _cotizacionRepository.Registrar(documentoBson);

            await _cotizacionRepository.Guardar(idBson, cotizacion.NroCotizacion, cotizacion.IdFormato, usuarioSesion);

            return new ResponseModel<string>(true, "Se ha guardado la cotización", idBson);
        }

        public async Task<ResponseModel<string>> ObtenerReporte(string codigoReporte)
        {

            CotizacionRegistroEntity datoReporte = await ObtenerDatosRegistro(codigoReporte);

            if (string.IsNullOrEmpty(datoReporte.Cotizacion))
                throw new NotFoundException("No se puedo encontrar el reporte");

            BsonDocument documento = await _cotizacionRepository.ObtenerDatosReporte(datoReporte.Codigo);

            string reporte = ReporteCotizacionFactory.GenerarReporte(datoReporte.IDFormato, documento);

            ResponseModel<string> response = new ResponseModel<string>(true, Constante.MESSSGE_SUCCESS_REPORT, reporte);

            return response;
        }

        public async Task<CotizacionAbstract> ObtenerDatosReporte(string codigoReporte)
        {
            CotizacionRegistroEntity datoReporte = await ObtenerDatosRegistro(codigoReporte);

            if (string.IsNullOrEmpty(datoReporte.Codigo))
                throw new NotFoundException("No se puedo encontrar el reporte");    

            BsonDocument documento = await _cotizacionRepository.ObtenerDatosReporte(codigoReporte);

            CotizacionAbstract cotizacionReporte = ReporteCotizacionFactory.ObtenerModeloCotizacion(datoReporte.IDFormato, documento);

            return cotizacionReporte;
        }

        public async Task<IEnumerable<FormatosPorClienteModel>> FormatosPorCliente(int idCliente)
        {
            IEnumerable<FormatosPorClienteModel> listaDeFormatos = await _cotizacionRepository.FormatosPorCliente(idCliente);
            return listaDeFormatos;
        }

        public async Task<IEnumerable<ReportesGeneradosPorCotizacionModel>> ReportesPorCotizacion(string cotizacion)
        {
            IEnumerable<ReportesGeneradosPorCotizacionModel> reportes = await _cotizacionRepository.ReportesPorCotizacion(cotizacion);
            return reportes;
        }

        private async Task<CotizacionRegistroEntity> ObtenerDatosRegistro(string codReporte)
        {
            CotizacionRegistroEntity result = await _cotizacionRepository.ObtenerDatosRegistro(codReporte);
            return result;
        }

        public async Task Actualizar(ActualizarReporteCotizacionModel reporte, int usuarioSesion)
        {
            BsonDocument documentoBson = BsonDocument.Parse(reporte.Cotizacion.ToString());
            await _cotizacionRepository.Actualizar(reporte.IdObject, usuarioSesion, documentoBson);
        }

    }
}
