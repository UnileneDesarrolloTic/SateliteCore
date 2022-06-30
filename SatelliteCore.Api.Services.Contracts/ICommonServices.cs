using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services.Contracts
{
    public interface ICommonServices
    {
        public Task<IEnumerable<TipoDocumentoIdentidadEntity>> ListarTipoDocumentoIndentidad();
        public Task<IEnumerable<PaisEntity>> ListarPaises();
        public Task<List<MenuxUsuarioModel>> ListarMenuxUsuario(int usuario);
        public Task<IEnumerable<RolEntity>> ListarRoles(string estado);
        public Task<List<FamiliaMP>> ListarFamiliaMP(string tipo);
        public Task<ResponseModel<IEnumerable<ConfiguracionEntity>>> ObtenerConfiguracionesSistema(int idConfiguracion, string grupo);
    }
}
