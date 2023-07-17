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

    }
}
