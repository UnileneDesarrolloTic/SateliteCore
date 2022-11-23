using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Generic;
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
        public Task<DatosClienteDTO> ObtenerDatosCliente(int codigoCliente);
        public Task ActualizarCorrelativoCodReclamo(int correlativo, int idConfiguracion, int id, string grupo);
        public Task<bool> ValidarExiteConfiguracionDetallePorId(int idConfiguracion, string estado = null);
        public Task<IEnumerable<TipoDocumentoSsomaEntity>> TipoDocumentoSsoma();
        public Task<IEnumerable<UbicacionSsomaEntity>> UbicacionSsoma();
        public Task<IEnumerable<ProteccionEntitySsoma>> ProteccionSsoma();
        public Task<IEnumerable<EstadoEntitySsoma>> EstadoSsoma();
        public Task<IEnumerable<AlmacenamientoSsomaEntity>> AlmacenamientoSsoma();
        public Task<IEnumerable<ResponsableSsomaEntity>> ResponsableSsoma();
    }
}
