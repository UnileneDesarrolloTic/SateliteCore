using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Contracts.Repository
{
    public interface IContabilidadRepository
    {
        public Task<IEnumerable<DetraccionesEntity>> ListarDetraccion();

        public int ProcesarDetraccionContabilidad(List<FormatoComprobantePagoDetraccion> dato);
    }
}
