using SatelliteCore.Api.Models.Dto.RRHH;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Contracts.Repository
{
    public interface IRRHHRepository
    {
        public Task<IEnumerable<ReporteAsistenciaDTO>> ListarAsistencia(DateTime fecha, int usuarioToken);
    }
}
