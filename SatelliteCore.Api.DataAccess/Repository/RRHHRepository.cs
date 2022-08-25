using Dapper;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Config;
using SatelliteCore.Api.Models.Dto.RRHH;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Repository
{
    public class RRHHRepository : IRRHHRepository
    {
        private readonly IAppConfig _appConfig;

        public RRHHRepository(IAppConfig appConfig)
        {
            _appConfig = appConfig;
        }

        public async Task<IEnumerable<ReporteAsistenciaDTO>> ListarAsistencia(DateTime fecha, int usuarioToken)
        {
            IEnumerable<ReporteAsistenciaDTO> listaAsistencia = new List<ReporteAsistenciaDTO>();

            using SqlConnection context = new SqlConnection(_appConfig.contextSpring);
            listaAsistencia = await context.QueryAsync<ReporteAsistenciaDTO>("usp_ObtenerReporteDiarioAsistencia_Satelite", new { fecha, usuario = usuarioToken }, commandType: CommandType.StoredProcedure);

            return listaAsistencia;
        }
    }
}
