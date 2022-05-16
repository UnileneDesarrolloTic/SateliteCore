using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services.Contracts
{
    public interface IContabilidadService
    {
        public Task<PaginacionModel<DetraccionesEntity>> ListarDetraccion( DatosListarDetraccionPaginado datos);
    }
}
