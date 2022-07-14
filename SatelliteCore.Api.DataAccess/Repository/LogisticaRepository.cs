using Dapper;
using SatelliteCore.Api.DataAccess.Contracts;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Config;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess
{
    public class LogisticaRepository : ILogisticaRepository
    {
        private readonly IAppConfig _appConfig;

        public LogisticaRepository(IAppConfig appConfig)
        {
            _appConfig = appConfig;
        }

        public bool RegistrarMaestroItem(DatosRequestMaestroItemModel dato)
        {
           
            return true;
        }
    }
}
