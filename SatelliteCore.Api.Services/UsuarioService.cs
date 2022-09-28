using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.CrossCutting.Helpers;
using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Services.Contracts;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services
{
    public class UsuarioService : IUsuarioService
    {
        private readonly IUsuarioRepository _usuarioRepository;

        public UsuarioService(IUsuarioRepository usuarioRepository)
        {
            _usuarioRepository = usuarioRepository;
        }

        public async Task<UsuarioEntity> ObtenerUsuario(DatoUsuarioBasico datos)
        {
            return await _usuarioRepository.ObtenerUsuario(datos);
        }

        public async Task<AuthResponse> ObtenerUsuarioLogin(AuthRequestModel datosUsuario)
        {
            datosUsuario.Clave = Encryptations.EncryptSha256(datosUsuario.Clave);
            return await _usuarioRepository.ObtenerUsuarioLogin(datosUsuario);
        }

        public async Task<int> CambiarClave(ActualizarClaveModel datos)
        {
            datos.Clave = Encryptations.EncryptSha256(datos.Clave);
            return await _usuarioRepository.CambiarClave(datos);
        }

        public async Task<PaginacionModel<UsuarioEntity>> ListarUsuarios(DatosListarUsuarioPaginado datos)
        {

            (List<UsuarioEntity> lista, int totalRegistros) resultDb = await _usuarioRepository.ListarUsuarios(datos);

            PaginacionModel<UsuarioEntity> response = new PaginacionModel<UsuarioEntity>(resultDb.lista, datos.Pagina, datos.RegistrosPorPagina, resultDb.totalRegistros);

            return response;
        }


        public async Task<DatosFormatoAsignacionPersonalLaboralModel> ListarAsignacionPersonal()
        {
            DatosFormatoAsignacionPersonalLaboralModel response = await _usuarioRepository.ListarAsignacionPersonal();
            return response;
        }

        public async Task<List<AreaPersonalLaboralEntity>> ListarAreaPersonaLaboral()
        {
            List<AreaPersonalLaboralEntity> response = await _usuarioRepository.ListarAreaPersonaLaboral();
            return response;    
        }

        public async Task<ResponseModel<string>> RegistrarPersonaLaboralMasiva(DatosFormatoAsignacionPersonalModel dato, int idUsuario)
        {
            int response = await _usuarioRepository.RegistrarPersonaLaboralMasiva(dato, idUsuario);

            return new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Registrado con exito");
        }

        public async Task<IEnumerable<DatosFormatoFiltrarTrabajadorAreaModel>> FiltrarAreaPersona(int idArea)
        {
            IEnumerable<DatosFormatoFiltrarTrabajadorAreaModel> response = await _usuarioRepository.FiltrarAreaPersona(idArea);
            return response;
        }

        public async Task<ResponseModel<string>> LiberalPersona(int IdAsignacion)
        {
            int response = await _usuarioRepository.LiberalPersona(IdAsignacion);

            return new ResponseModel<string>(true, Constante.MESSAGE_SUCCESS, "Liberado con exito");
        }


    }
}
