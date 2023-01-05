using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Dto.RRHH;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Services.Contracts;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services
{
    public class RRHHServices : IRRHHServices
    {
        private readonly IRRHHRepository _rrhhRepository;
        public RRHHServices(IRRHHRepository rrhhRepository)
        {
            _rrhhRepository = rrhhRepository;
        }

        public async Task<ResponseModel<IEnumerable<ReporteAsistenciaDTO>>> ListarAsistencia(DateTime fecha, int usuarioToken)
        {
            IEnumerable<ReporteAsistenciaDTO> listaAsistencia = await _rrhhRepository.ListarAsistencia(fecha, usuarioToken);
            return new ResponseModel<IEnumerable<ReporteAsistenciaDTO>>(true, Constante.MESSAGE_SUCCESS, listaAsistencia);
        }

        public async Task<ResponseModel<string>> RegistrarHorasExtras(DatosEstructuraHorasExtraCabecera data, string usuario)
        {
            int result = 0;
            result=await _rrhhRepository.RegistrarHorasExtras(data, usuario);
            return new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Registrado con exito");
        }

        public async Task<IEnumerable<DatosFormatoListarHorasExtrasPersona>> ListarHoraExtrasPersona(DatosFormatoFiltraHoraExtras data)
        {
            IEnumerable<DatosFormatoListarHorasExtrasPersona> response = new List<DatosFormatoListarHorasExtrasPersona>();
            response = await _rrhhRepository.ListarHoraExtrasPersona(data);
            return response;
        }

        public async Task<object> BuscarInformacionHorasExtrasPersona(int Cabecera)
        {
            (DatosFormatoHorasExtrasCabeceraModel cabecera, List <DatosFormatoHorasExtrasDetalle> detalle) = await _rrhhRepository.BuscarInformacionHorasExtrasPersona(Cabecera);
            object result = new { cabecera, detalle };

            return result;
        }
    }
}
