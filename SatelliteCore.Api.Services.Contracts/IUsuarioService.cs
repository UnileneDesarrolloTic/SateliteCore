﻿using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services.Contracts
{
    public interface IUsuarioService
    {
        public Task<AuthResponse> ObtenerUsuarioLogin(AuthRequestModel datosUsuario);
        public Task<UsuarioEntity> ObtenerUsuario(DatoUsuarioBasico datos);
        public Task<PaginacionModel<UsuarioEntity>> ListarUsuarios(DatosListarUsuarioPaginado datos);
        public Task<int> CambiarClave(ActualizarClaveModel datos);
        public Task<DatosFormatoAsignacionPersonalLaboralModel> ListarAsignacionPersonal();
        public Task<List<AreaPersonalLaboralEntity>> ListarAreaPersonaLaboral();
        public Task<ResponseModel<string>> RegistrarPersonaLaboralMasiva(DatosFormatoAsignacionPersonalModel dato,int idUsuario);
        public Task<IEnumerable<DatosFormatoFiltrarTrabajadorAreaModel>> FiltrarAreaPersona(int idArea);
        public Task<ResponseModel<string>> LiberalPersona(int IdAsignacion);
    }
}
