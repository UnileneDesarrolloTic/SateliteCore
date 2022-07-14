using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.DataAccess.Contracts;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Services.Contracts;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services
{
    public class LogisticaServices : ILogisticaServices
    {
        private readonly ILogisticaRepository _LogisticaRepository;

        public LogisticaServices(ILogisticaRepository LogisticaRepository)
        {
            _LogisticaRepository = LogisticaRepository;
        }

        public bool RegistrarMaestroItem(DatosRequestMaestroItemModel dato)
        {
            return _LogisticaRepository.RegistrarMaestroItem(dato);
        }



    }
}
