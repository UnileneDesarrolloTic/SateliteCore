using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;
using System.Threading.Tasks;


namespace SatelliteCore.Api.Services.Contracts
{
    public interface IContabilidadService
    {
        public Task<List<DetraccionesEntity>> ListarDetraccion();

        public Task<int>ProcesarDetraccionContabilidad(List<FormatoComprobantePagoDetraccion> dato);

        
    }
}
