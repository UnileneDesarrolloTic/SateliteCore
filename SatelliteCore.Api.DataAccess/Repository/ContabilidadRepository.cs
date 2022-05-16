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
    public class ContabilidadRepository : IContabilidadRepository
    {
        private readonly IAppConfig _appConfig;

        public ContabilidadRepository(IAppConfig appConfig)
        {
            _appConfig = appConfig;
        }


        public async Task<(List<DetraccionesEntity>, int)> ListarDetraccion(DatosListarDetraccionPaginado datos)
        {
            (List<DetraccionesEntity> ListaUsuarios, int totalRegistros) result;

            using (var connection = new SqlConnection(_appConfig.contextSatelliteDB))
            {
                using (var result_db = await connection.QueryMultipleAsync("usp_ListarDetraccionContabilidad", datos , commandType: CommandType.StoredProcedure))
                {
                    result.ListaUsuarios = result_db.Read<DetraccionesEntity>().ToList();
                    result.totalRegistros = result_db.Read<int>().First();
                }

                connection.Dispose();
            }

            return result;
        }
    }
}
