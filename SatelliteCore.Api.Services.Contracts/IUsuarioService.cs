using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Request.AsignacionPersonal;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.RRHH.AsignacionPersonal;
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
        public Task<IEnumerable<DatosFormatoFiltrarTrabajadorAreaModel>> FiltrarAreaPersona(int idArea,string NombrePersona);
        public Task<ResponseModel<string>> LiberalPersona(int IdAsignacion);
        public Task<ResponseModel<string>> ExportarExcelPersonaAsignacion(DatosFormatoFiltroAsignacionPersona dato);
        public Task<ResponseModel<AreaPersonalLaboralEntity>> RegistrarEditarArea(int IdArea, string Descripcion);
        public Task<ResponseModel<string>> EliminarAreaProduccion(int IdArea);
        public Task<IEnumerable<DatosFormatoListarPersonaTecnica>> ListarPersonaTecnico();
        public Task<IEnumerable<DatosFormatosPersonaPorAreaModel>> ListarPersonaPorArea(int IdArea);
        public Task<IEnumerable<DatosFormatoPersonasAsistencia>> MostrarPersonasAsistencias(string idArea);
    }
}
