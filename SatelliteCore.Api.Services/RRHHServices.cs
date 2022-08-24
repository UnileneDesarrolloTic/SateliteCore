using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Dto.RRHH;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Services.Contracts;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services
{
    public class RRHHServices : IRRHHServices
    {
        private readonly IRRHHRepository _rrhhRepository;
        public RRHHServices(IRRHHRepository rrhhRepository)
        {
            _rrhhRepository = rrhhRepository;
        }

        public async Task<ResponseModel<IEnumerable<ReporteAsistenciaDTO>>> ListarAsistencia(DateTime fecha, int usuarioToken)
        {
            IEnumerable<ReporteAsistenciaDTO> listaAsistencia = await _rrhhRepository.ListarAsistencia(fecha, usuarioToken);
            return new ResponseModel<IEnumerable<ReporteAsistenciaDTO>>(true, Constante.MESSAGE_SUCCESS, listaAsistencia);
        }
    }
}
