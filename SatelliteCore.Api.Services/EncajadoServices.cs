using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Encajado;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Services.Contracts;
using System.Collections.Generic;
using System.Threading.Tasks;
using SystemsIntegration.Api.Models.Exceptions;

namespace SatelliteCore.Api.Services
{
    public class EncajadoServices : IEncajadoServices
    {
        private readonly IEncajadoRespository _encajadoRespository;

        public EncajadoServices(IEncajadoRespository encajadoRespository)
        {
            _encajadoRespository = encajadoRespository;
        }

        public async Task<ResponseModel<List<ListaOrdenesFabricaciónDTO>>> ListaOrdenesFabricacion(string ordenFabricacion)
        {

            List<ListaOrdenesFabricaciónDTO> lista = await _encajadoRespository.ListaOrdenesFabricacion(ordenFabricacion);

            if (lista.Count < 1)
                return new ResponseModel<List<ListaOrdenesFabricaciónDTO>>(false, "No se han encontrado registros.", lista);

            return new ResponseModel<List<ListaOrdenesFabricaciónDTO>>(lista); 
        }
    }
}
