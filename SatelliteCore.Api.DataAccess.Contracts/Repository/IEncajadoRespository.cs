using SatelliteCore.Api.Models.Encajado;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Contracts.Repository
{
    public interface IEncajadoRespository
    {
        public Task<List<ListaOrdenesFabricaciónDTO>> ListaOrdenesFabricacion(string ordenFabricacion);
    }
}
