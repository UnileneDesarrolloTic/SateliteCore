using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.RRHH.AsignacionPersonal;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Contracts.Repository
{
    public interface IUsuarioRepository
    {
        public Task<AuthResponse> ObtenerUsuarioLogin(AuthRequestModel datosUsuario);
        public Task<UsuarioEntity> ObtenerUsuario(DatoUsuarioBasico datos);
        public Task<(List<UsuarioEntity>, int)> ListarUsuarios(DatosListarUsuarioPaginado datos);
        public Task<int> CambiarClave(ActualizarClaveModel datos);
        public Task<DatosFormatoAsignacionPersonalLaboralModel> ListarAsignacionPersonal();
        public Task<List<AreaPersonalLaboralEntity>> ListarAreaPersonaLaboral();
        public Task<int> RegistrarPersonaLaboralMasiva(DatosFormatoAsignacionPersonalModel dato, int idUsuario);
        public Task<IEnumerable<DatosFormatoFiltrarTrabajadorAreaModel>> FiltrarAreaPersona(int idArea,string NombrePersona);
        public Task<int> LiberalPersona(int IdAsignacion);
        public Task<IEnumerable<DatosFormatoPersonaAsignacionExportModel>> ExportarExcelPersonaAsignacion(string FechaInicio, string FechaFinal);
        public Task<AreaPersonalLaboralEntity> RegistrarEditarArea(int IdArea, string Descripcion);
        public Task<int> EliminarAreaProduccion(int IdArea);
        public Task<IEnumerable<DatosFormatoListarPersonaTecnica>> ListarPersonaTecnico();
        public Task<IEnumerable<DatosFormatosPersonaPorAreaModel>> ListarPersonaPorArea(int IdArea);
        public Task<IEnumerable<DatosFormatoPersonasAsistencia>> MostrarPersonasAsistencias(string idArea);
    }
}
