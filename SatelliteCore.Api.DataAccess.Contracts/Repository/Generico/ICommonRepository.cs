using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
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
        public Task<IEnumerable<FamiliaMaestroItemsModel>> ListarFamiliaGeneral(string idlinea);
        public Task<IEnumerable<SubFamiliaEntity>> ListarSubFamilia(string idlinea,string idfamilia);
        public Task<IEnumerable<MarcaEntity>> ListarMarca();
        public Task<FormatoResponseRegistrarMaestroItem> RegistrarMaestroItem(DatosRequestMaestroItemModel dato,int idUsuario);
        public Task<(List<FormatoListarMaestroItemModel>, int)> ListarMaestroItem(DatosListarMaestroItemPaginador datos);
        public Task<IEnumerable<MaestroAlmacenEntity>> ListarMaestroAlmacen();
        public Task<bool> ValidacionPermisoAccesso(string Permiso, int idUsuario);
        public Task<IEnumerable<GrupoEntity>> ListarGrupo();
        public Task<IEnumerable<TablaEntity>> ListarTabla(string Grupo);
        public Task<IEnumerable<MarcaProtocoloEntity>> ListarMarcaProtocolo(string Grupo,string Campo);
        public Task<IEnumerable<MetodologiaEntity>> ListarMetodologiaProtocolo();
    }
}
