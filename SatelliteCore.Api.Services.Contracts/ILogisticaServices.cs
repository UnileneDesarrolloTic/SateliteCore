using SatelliteCore.Api.Models.Dto.AnalisisAgujas;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services.Contracts
{
    public interface ILogisticaServices
    {
        public Task<IEnumerable<DatosFormatoPlanOrdenServicosD>> ObtenerNumeroGuias(string NumeroGuia);

        public Task<ResponseModel<string>> RegistrarRetornoGuia(List<DatosFormatoRetornoGuiaRequest> dato);
    }
}
