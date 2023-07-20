using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Config;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Services.Contracts;

namespace SatelliteCore.Api.Services
{
    public class RegistroAsistenciaServices : IRegistroAsistenciaServices
    {
        private readonly IRegistroAsistenciaRepository _registroAsistenciaRepository;



        public RegistroAsistenciaServices(IRegistroAsistenciaRepository registroAsistenciaRepository)
        {
            _registroAsistenciaRepository = registroAsistenciaRepository;
        }

        public async Task<ResponseModel<string>> RegistraAsistencia(string numeroDocumento)
        {
            bool resultado;
            resultado = await _registroAsistenciaRepository.RegistraAsistencia(numeroDocumento);

            if (resultado)
                return new ResponseModel<string>(true, "Se ha guardado su asistencia", "");

            return new ResponseModel<string>(false, "El numero de identificación no se encuentra en nuestro sistema", "");
        }
    }
}
