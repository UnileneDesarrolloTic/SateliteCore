using MongoDB.Bson;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Contracts.Repository
{
    public interface ICotizacionRepository
    {
        public Task<(List<CotizacionEntity>, int)> Listar(DatosListarCotizacionesPaginado datos);
        public Task<ObtenerEstructuraFormCotizacionModel> FormatoEstructura(int codFormato);
        public Task<(object cabecera, object detalle)> FormatoDatos(int idFormato, string cotizacion);
        public Task<string> Registrar(BsonDocument cotizacion);
        public Task Guardar(string id, string cotizacion, int idFormato, int usuarioSesion);
        public Task<IEnumerable<FormatosPorClienteModel>> FormatosPorCliente(int idCliente);
        public Task<IEnumerable<ReportesGeneradosPorCotizacionModel>> ReportesPorCotizacion(string cotizacion);
        public Task<BsonDocument> ObtenerDatosReporte(string idObject);
        public Task<CotizacionRegistroEntity> ObtenerDatosRegistro(string codReporte);
        public Task Actualizar(string id, int usuarioSesion, BsonDocument cotizacion);
    }
}
