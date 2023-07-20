using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Config;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;


namespace SatelliteCore.Api.DataAccess.Repository
{
    public class RegistroAsistenciaRepository : IRegistroAsistenciaRepository
    {
        private readonly IAppConfig _appConfig;

        public RegistroAsistenciaRepository(IAppConfig appConfig)
        {
            _appConfig = appConfig;
        }

        public async Task<bool> RegistraAsistencia(string numeroDocumento)
        {
            bool result;

            using (var satelliteContext = new SqlConnection(_appConfig.contextSpring))
            {
                result = await satelliteContext.QuerySingleAsync<bool>("usp_Satelite_RegistroAsistencia", new { numeroDocumento }, commandType: CommandType.StoredProcedure);
                satelliteContext.Dispose();
            }

            return result;
        }
    }
}
