using SatelliteCore.Api.Models.Dto.AnalisisAgujas;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Contracts.Repository
{
    public interface IAnalisisAgujaRepository
    {
        public Task<IEnumerable<ListarAnalisisAgujaModel>> ListarAnalisisAguja(string ordenCompra, string lote, string item, int pagina);
        public Task<IEnumerable<ListarOrdenCompra>> ListaOrdenesCompra(string NumeroOrden);
        public Task<int> CantidadPruebaFlexionPorItem(string controlNumero, int secuencia);
        public Task<string> RegistrarAnalisisAguja(ControlAgujasModel matricula);
        public Task ValidarLoteCreado(string controlNumero, int secuencia, int codUsuarioSesion);
        public Task<(ObtenerAnalisisAgujaModel, List<AnalisisAgujaFlexionEntity>)> AnalisisAgujaFlexion(string loteAnalisis);
        public Task EliminarPruebaFlexionAguja(string loteAnalisis);
        public Task GuardarPruebaFlexionAguja(DatosFormatoRegistroPruebasAgujasModel dato);
        public Task<ObtenerDatosGeneralesDTO> ObtenerDatosGenerales(string loteAnalisis);
        public Task<AnalisisAgujaPlanMuestreoEntity> ObtenerPlanMuestreo(string loteAnalisis);
        public Task RegistrarPlanMuestreo(AnalisisAgujaPlanMuestreoEntity planMuestreo);
        public Task EliminarPlanMuestreo(string loteAnalisis);
        public Task<List<AnalisisAgujaPruebaDimensionalEntity>> ObtenerPruebaDimensional(string loteAnalisis);
        public Task EliminarPruebaDimensional(string loteAnalisis);
        public Task RegistrarPruebaDimensional(List<AnalisisAgujaPruebaDimensionalEntity> prueba);
        public Task<IEnumerable<AnalisisAgujaElasticidadPerforacionEntity>> ObtenerPruebaElasticidadPerforacion(string loteAnalisis);
        public Task EliminarPruebaElasticidadPerforacion(string loteAnalisis);
        public Task RegistrarPruebaElasticidadPerforacion(List<AnalisisAgujaElasticidadPerforacionEntity> prueba);
        public Task<IEnumerable<AnalisisAgujaPruebaAspectoEntity>> ObtenerPruebaAspecto(string loteAnalisis);
        public Task EliminarPruebaAspecto(string loteAnalisis);
        public Task RegistrarPruebaAspecto(PruebaAspectoYObservacionesDTO datos, string loteAnalisis);
    }
}
