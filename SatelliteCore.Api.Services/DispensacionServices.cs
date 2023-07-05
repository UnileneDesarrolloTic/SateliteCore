using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Services.Contracts;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using SystemsIntegration.Api.Models.Exceptions;

namespace SatelliteCore.Api.Services
{
    public class DispensacionServices : IDispensacionServices
    {
        private readonly IDispensacionRepository _dispensacionRepository;

        public DispensacionServices(IDispensacionRepository dispensacionRepository)
        {
            _dispensacionRepository = dispensacionRepository;
           
        }
    }
}
