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
        private readonly ILogisticaRepository _logisticaRepository;

        public LogisticaServices(ILogisticaRepository logisticaRepository)
        {
            _logisticaRepository = logisticaRepository;
        }
        public async Task<IEnumerable<DatosFormatoPlanOrdenServicosD>> ObtenerNumeroGuias(string NumeroGuia)
        {
            return await _logisticaRepository.ObtenerNumeroGuias(NumeroGuia);
        }


    }
}
