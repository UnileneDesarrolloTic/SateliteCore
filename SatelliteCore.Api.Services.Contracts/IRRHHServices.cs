using SatelliteCore.Api.Models.Dto.RRHH;
using SatelliteCore.Api.Models.Response;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services.Contracts
{
    public interface IRRHHServices
    {
        public Task<ResponseModel<IEnumerable<ReporteAsistenciaDTO>>> ListarAsistencia(DateTime fecha, int usuarioToken);
    }
}
