using SatelliteCore.Api.Models.Entities;
using System.Collections.Generic;
using System.Threading.Tasks;


namespace SatelliteCore.Api.Services.Contracts
{
    public interface IContabilidadService
    {
        public Task<IEnumerable<DetraccionesEntity>> ListarDetraccion();

        public int ProcesarDetraccionContabilidad(string urlarchivo);

        
    }
}
