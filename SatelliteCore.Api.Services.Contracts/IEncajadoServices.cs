using SatelliteCore.Api.Models.Encajado;
using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services.Contracts
{
    public interface IEncajadoServices
    {
        public Task<ResponseModel<List<ListaOrdenesFabricaciónDTO>>> ListaOrdenesFabricacion(string ordenFabricacion);
    }
}
