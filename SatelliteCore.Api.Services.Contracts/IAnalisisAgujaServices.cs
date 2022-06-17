using SatelliteCore.Api.Models.Dto.AnalisisAgujas;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services.Contracts
{
    public interface IAnalisisAgujaServices
    {
        public Task<IEnumerable<ListarAnalisisAgujaModel>> ListarAnalisisAguja(string ordenCompra, string lote, string item, int pagina);
        public Task<IEnumerable<ListarOrdenCompra>> ListaOrdenesCompra(string numeroOrden);
        public Task<ResponseModel<int>> CantidadPruebaFlexionPorItem(string controlNumero, int secuencia);
        public Task<ResponseModel<object>> RegistrarAnalisisAguja(ControlAgujasModel matricula);
        public Task<ResponseModel<string>> ValidarLoteCreado(string controlNumero, int secuencia, int codUsuarioSesion);
        public Task<object> AnalisisAgujaFlexion(string loteAnalisis);
        public Task<ResponseModel<string>> GuardarEditarPruebaFlexionAguja(List<GuardarPruebaFlexionAgujaModel> analisis);
        public Task<ResponseModel<string>> ReporteAnalisisFlexion(string loteAnalisis);
        public Task<ResponseModel<ObtenerDatosGeneralesDTO>> ObtenerDatosGenerales(string loteAnalisis);
        public Task<ResponseModel<AnalisisAgujaPlanMuestreoEntity>> ObtenerPlanMuestreo(string loteAnalisis);
        public Task<ResponseModel<string>> GuardarPlanMuestreo(AnalisisAgujaPlanMuestreoEntity planMuestreo);
        public Task<ResponseModel<List<AnalisisAgujaPruebaDimensionalEntity>>> ObtenerPruebaDimensional(string loteAnalisis);
        public Task<ResponseModel<string>> GuardarPruebaDimensional(List<AnalisisAgujaPruebaDimensionalEntity> prueba);
        public Task<ResponseModel<List<AnalisisAgujaElasticidadPerforacionEntity>>> ObtenerPruebaElasticidadPerforacion(string loteAnalisis);
        public Task<ResponseModel<string>> GuardarPruebaElasticidadPerforacion(List<AnalisisAgujaElasticidadPerforacionEntity> prueba);
        public Task<ResponseModel<List<AnalisisAgujaPruebaAspectoEntity>>> ObtenerPruebaAspecto(string loteAnalisis);
        public Task<ResponseModel<string>> GuardarPruebaAspecto(PruebaAspectoYObservacionesDTO datos, int usuario);
        public Task<ResponseModel<string>> ObtenerReporteAnalisisAguja(string loteAnalisis);
    }
}
