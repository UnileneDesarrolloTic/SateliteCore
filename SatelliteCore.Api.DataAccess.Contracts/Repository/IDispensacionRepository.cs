using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Request.Dispensacion;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.Dispensacion;
using System.Collections.Generic;
using System.Threading.Tasks;


namespace SatelliteCore.Api.DataAccess.Contracts.Repository
{
    public interface IDispensacionRepository
    {
        public Task<IEnumerable<DatosFormatoObtenerOrdenFabricacion>> ObtenerOrdenFabricacion(DatosFormatoFiltroOrdenFabricacion dato);
        public Task<IEnumerable<DatosFormatoListadoMateriaPrimaDispensacion>> RecetasOrdenFabricacion(string ordenFabricacion);
        public Task<string> RegistrarDispensacionMP(DatosFormatoDispensacionMateriaPrima dato, string usuario);
        public Task<IEnumerable<DatosFormatoHistorialDispensaccion>> HistorialDispensacionMP(string ordenFabricacion, string lote);
        public Task<DatosFormatoInformacionDispensacionPT> InformacionItem(string item, string ordenFabricacion, string secuencia);
        public Task<DatosFormatoDispensacionDetalle> DetalleDispensacionReceta();
        public Task<string> RegistrarRecetasGlobal(IEnumerable<DatosFormatoRegistroDispensacionRecetaGlobal> dato, string usuario);
        public Task<IEnumerable<DatosFormatoDispensacionGuiaDespacho>> DispensacionGuiaDespacho(DatosFormatoFiltroDispensacion dato);
        public Task<IEnumerable<DatosFormatoMostrarDispensacionDespacho>> MostrarDispensacionDespacho(string id);
    }
}
