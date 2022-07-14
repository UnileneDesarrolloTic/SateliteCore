using SatelliteCore.Api.Models.Entities;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Contracts.Repository
{
    public interface ICommonRepository
    {
        public Task<IEnumerable<TipoDocumentoIdentidadEntity>> ListarTipoDocumentoIndentidad();
        public Task<IEnumerable<PaisEntity>> ListarPaises();
        public Task<List<MenuEntity>> ListarMenuxUsuario(int usuario);
        public Task<IEnumerable<RolEntity>> ListarRoles(string estado);
        public Task<List<FamiliaMP>> ListarFamiliaMP(string tipo);
        public Task<IEnumerable<ConfiguracionEntity>> ObtenerConfiguracionesSistema(int idConfiguracion, string grupo);
        public Task<IEnumerable<AgrupadorEntity>> ListarAgrupador();
        public Task<IEnumerable<SubAgrupadorEntity>> ListarSubAgrupador(string idAgrupador);
        public Task<IEnumerable<LineaEntity>> ListarLinea();

        public Task<IEnumerable<FamiliaMaestroItemsModel>> ListarFamilia(string idlinea);

        public Task<IEnumerable<SubFamiliaEntity>> ListarSubFamilia(string idlinea,string idfamilia);
    }
}
