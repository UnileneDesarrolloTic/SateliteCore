using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using SatelliteCore.Api.Models.Response;

namespace SatelliteCore.Api.Services.Contracts
{
    public interface IRegistroAsistenciaServices
    {
        public Task<ResponseModel<string>> RegistraAsistencia(string numeroDocumento);

    }
}
