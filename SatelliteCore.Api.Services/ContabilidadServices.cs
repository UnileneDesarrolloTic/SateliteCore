using SatelliteCore.Api.CrossCutting.Helpers;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Services.Contracts;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services
{
    public class ContabilidadServices : IContabilidadService
    {
        private readonly IContabilidadRepository _contabilidadRepository;

        public ContabilidadServices(IContabilidadRepository contabilidadRepository)
        {
            _contabilidadRepository = contabilidadRepository;
        }
        public async Task<PaginacionModel<DetraccionesEntity>> ListarDetraccion(DatosListarDetraccionPaginado datos)
        {

            (List<DetraccionesEntity> lista, int totalRegistros) resultDb = await _contabilidadRepository.ListarDetraccion(datos);

            PaginacionModel<DetraccionesEntity> response = new PaginacionModel<DetraccionesEntity>(resultDb.lista, datos.Pagina, datos.RegistrosPorPagina, resultDb.totalRegistros);

            return response;
        }

    }
}
