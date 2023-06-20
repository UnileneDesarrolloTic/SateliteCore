using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
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
        public Task<ResponseModel<IEnumerable<AgrupadorEntity>>> ListarAgrupador();
        public Task<ResponseModel<IEnumerable<SubAgrupadorEntity>>> ListarSubAgrupador(string idAgrupador);
        public Task<ResponseModel<IEnumerable<LineaEntity>>> ListarLinea();
        public Task<ResponseModel<IEnumerable<FamiliaMaestroItemsModel>>> ListarFamilia(string idlinea);
        public Task<ResponseModel<IEnumerable<FamiliaMaestroItemsModel>>> ListarFamiliaGeneral(string idlinea);
        public Task<ResponseModel<IEnumerable<SubFamiliaEntity>>> ListarSubFamilia(string idlinea,string idfamilia);
        public Task<ResponseModel<IEnumerable<MarcaEntity>>> ListarMarca();
        public Task<ResponseModel<object>> RegistrarMaestroItem(DatosRequestMaestroItemModel dato, int idUsuario);
        public Task<PaginacionModel<FormatoListarMaestroItemModel>> ListarMaestroItem(DatosListarMaestroItemPaginador datos);
        public Task<ResponseModel<IEnumerable<MaestroAlmacenEntity>>> ListarMaestroAlmacen();
        public Task<ResponseModel<bool>> ValidacionPermisoAccesso(string Permiso, int idUsuario);
        public Task<ResponseModel<IEnumerable<GrupoEntity>>> ListarGrupo();
        public Task<ResponseModel<IEnumerable<TablaEntity>>> ListarTabla(string Grupo);
        public Task<ResponseModel<IEnumerable<MarcaProtocoloEntity>>> ListarMarcaProtocolo(string Grupo, string Campo);
        public Task<ResponseModel<IEnumerable<MetodologiaEntity>>> ListarMetodologiaProtocolo();
        public Task<ResponseModel<IEnumerable<AgrupadorHebrasEntity>>> ListarAgrupadoHebra();
        public Task<ResponseModel<IEnumerable<CalibrePruebaEntity>>> ListarCalibrePrueba();
        public Task<ResponseModel<IEnumerable<TipoDocumentoSsomaEntity>>> TipoDocumentoSsoma();
        public Task<ResponseModel<IEnumerable<UbicacionSsomaEntity>>> UbicacionSsoma();
        public Task<ResponseModel<IEnumerable<ProteccionEntitySsoma>>> ProteccionSsoma();
        public Task<ResponseModel<IEnumerable<EstadoEntitySsoma>>> EstadoSsoma();
        public Task<ResponseModel<IEnumerable<AlmacenamientoSsomaEntity>>> AlmacenamientoSsoma();
        public Task<ResponseModel<IEnumerable<ResponsableSsomaEntity>>> ResponsableSsoma();
        public Task<ResponseModel<DatosClienteDTO>> ObtenerDatosCliente(int codigoCliente);
        public Task<ResponseModel<IEnumerable<TransportistaEntity>>> Transportista();
        public Task<ResponseModel<IEnumerable<ClasificacionAreaEntity>>> ClasificacionArea();

    }
}
