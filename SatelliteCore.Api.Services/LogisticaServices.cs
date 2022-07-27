using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Dto.AnalisisAgujas;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.ReportServices.Contracts.AnalsisAguja;
using SatelliteCore.Api.Services.Contracts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using SystemsIntegration.Api.Models.Exceptions;

namespace SatelliteCore.Api.Services
{
    public class LogisticaServices : ILogisticaServices
    {
        private readonly ILogisticaServices _logisticaRepository;

        public LogisticaServices(ILogisticaServices logisticaRepository)
        {
            _logisticaRepository = logisticaRepository;
        }


    }
}
