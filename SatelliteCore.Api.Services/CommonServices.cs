using SatelliteCore.Api.CrossCutting.Config;
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
    public class CommonServices : ICommonServices
    {
        private readonly ICommonRepository _commonRepository;

        public CommonServices(ICommonRepository commonRepository)
        {
            _commonRepository = commonRepository;
        }

        public async Task<IEnumerable<TipoDocumentoIdentidadEntity>> ListarTipoDocumentoIndentidad()
        {
            return await _commonRepository.ListarTipoDocumentoIndentidad();
        }

        public async Task<IEnumerable<PaisEntity>> ListarPaises()
        {
            return await _commonRepository.ListarPaises();
        }

        public async Task<List<MenuxUsuarioModel>> ListarMenuxUsuario(int usuario)
        {
            List<MenuEntity> listaMenu = await _commonRepository.ListarMenuxUsuario(usuario);

            List<MenuxUsuarioModel> menuResponse = new List<MenuxUsuarioModel>();

            int index = -1;

            foreach (MenuEntity menu in listaMenu)
            {
                if (menu.MenuPadre == 0)
                {
                    menuResponse.Add(new MenuxUsuarioModel()
                    {
                        Codigo = menu.Codigo,
                        Ruta = menu.Ruta,
                        Titulo = menu.Titulo,
                        Icono = menu.Icono,
                        Clase = menu.Clase,
                        ClaseEtiqueta = menu.ClaseEtiqueta,
                        ExtraLink = menu.ExtraLink,
                        MenuPadre = menu.MenuPadre,
                        subMenu = new List<MenuxUsuarioModel>()
                    });
                }
                else
                {
                    index = menuResponse.FindIndex(x => x.Codigo == menu.MenuPadre && x.MenuPadre == 0);

                    if (index != -1)
                    {
                        menuResponse.Add(new MenuxUsuarioModel()
                        {
                            Codigo = menu.Codigo,
                            Ruta = menu.Ruta,
                            Titulo = menu.Titulo,
                            Icono = menu.Icono,
                            Clase = menu.Clase,
                            ClaseEtiqueta = menu.ClaseEtiqueta,
                            ExtraLink = menu.ExtraLink,
                            MenuPadre = menu.MenuPadre,
                            subMenu = new List<MenuxUsuarioModel>()
                        });
                    }
                    else
                    {
                        index = -1;

                        index = menuResponse.FindIndex(x => x.Codigo == menu.MenuPadre && x.MenuPadre != 0);

                        if (index != -1)
                        {
                            menuResponse[index].subMenu.Add(new MenuxUsuarioModel()
                            {
                                Codigo = menu.Codigo,
                                Ruta = menu.Ruta,
                                Titulo = menu.Titulo,
                                Icono = menu.Icono,
                                Clase = menu.Clase,
                                ClaseEtiqueta = menu.ClaseEtiqueta,
                                ExtraLink = menu.ExtraLink,
                                MenuPadre = menu.MenuPadre,
                                subMenu = new List<MenuxUsuarioModel>()
                            });
                        }

                        else
                        {
                            menuResponse.Add(new MenuxUsuarioModel()
                            {
                                Codigo = menu.Codigo,
                                Ruta = menu.Ruta,
                                Titulo = menu.Titulo,
                                Icono = menu.Icono,
                                Clase = menu.Clase,
                                ClaseEtiqueta = menu.ClaseEtiqueta,
                                ExtraLink = menu.ExtraLink,
                                MenuPadre = menu.MenuPadre,
                                subMenu = new List<MenuxUsuarioModel>()
                            });
                        }
                    }

                    index = -1;
                }
            }

            return menuResponse;
        }

        public async Task<IEnumerable<RolEntity>> ListarRoles(string estado)
        {
            return await _commonRepository.ListarRoles(estado);
        }

        public async Task<List<FamiliaMP>> ListarFamiliaMP(string tipo)
        {
            List<FamiliaMP> familia = new List<FamiliaMP>();
            familia = await _commonRepository.ListarFamiliaMP(tipo);
            if (familia.Count>0)
                familia.Insert(0,new FamiliaMP  { Codigo="TD" , Valor1="Todos" });

            return familia;
        }


        public async Task<ResponseModel<IEnumerable<ConfiguracionEntity>>> ObtenerConfiguracionesSistema(int idConfiguracion, string grupo)
        {
            IEnumerable<ConfiguracionEntity> configuraciones = await _commonRepository.ObtenerConfiguracionesSistema(idConfiguracion, grupo);
            ResponseModel<IEnumerable<ConfiguracionEntity>> resultadoConfiguraciones = new ResponseModel<IEnumerable<ConfiguracionEntity>>(true, Constante.MESSAGE_SUCCESS, configuraciones);

            return resultadoConfiguraciones;
        }


        public async Task<ResponseModel<IEnumerable<AgrupadorEntity>>> ListarAgrupador()
        {
            IEnumerable<AgrupadorEntity> configuraciones = await _commonRepository.ListarAgrupador();
            ResponseModel<IEnumerable<AgrupadorEntity>> resultadoConfiguraciones = new ResponseModel<IEnumerable<AgrupadorEntity>>(true, Constante.MESSAGE_SUCCESS, configuraciones);

            return resultadoConfiguraciones;
        }

        public async Task<ResponseModel<IEnumerable<SubAgrupadorEntity>>> ListarSubAgrupador(string idAgrupador)
        {
            IEnumerable<SubAgrupadorEntity> configuraciones = await _commonRepository.ListarSubAgrupador(idAgrupador);
            ResponseModel<IEnumerable<SubAgrupadorEntity>> resultadoConfiguraciones = new ResponseModel<IEnumerable<SubAgrupadorEntity>>(true, Constante.MESSAGE_SUCCESS, configuraciones);

            return resultadoConfiguraciones;
        }

        public async Task<ResponseModel<IEnumerable<LineaEntity>>> ListarLinea()
        {
            IEnumerable<LineaEntity> configuraciones = await _commonRepository.ListarLinea();
            ResponseModel<IEnumerable<LineaEntity>> resultadoConfiguraciones = new ResponseModel<IEnumerable<LineaEntity>>(true, Constante.MESSAGE_SUCCESS, configuraciones);

            return resultadoConfiguraciones;
        }

        public async Task<ResponseModel<IEnumerable<FamiliaMaestroItemsModel>>> ListarFamilia(string idlinea)
        {
            IEnumerable<FamiliaMaestroItemsModel> configuraciones = await _commonRepository.ListarFamilia(idlinea);
            ResponseModel<IEnumerable<FamiliaMaestroItemsModel>> resultadoConfiguraciones = new ResponseModel<IEnumerable<FamiliaMaestroItemsModel>>(true, Constante.MESSAGE_SUCCESS, configuraciones);

            return resultadoConfiguraciones;
        }

        public async Task<ResponseModel<IEnumerable<FamiliaMaestroItemsModel>>> ListarFamiliaGeneral(string idlinea)
        {
            IEnumerable<FamiliaMaestroItemsModel> configuraciones = await _commonRepository.ListarFamiliaGeneral(idlinea);
            ResponseModel<IEnumerable<FamiliaMaestroItemsModel>> resultadoConfiguraciones = new ResponseModel<IEnumerable<FamiliaMaestroItemsModel>>(true, Constante.MESSAGE_SUCCESS, configuraciones);

            return resultadoConfiguraciones;
        }

        public async Task<ResponseModel<IEnumerable<SubFamiliaEntity>>> ListarSubFamilia(string idlinea, string idfamilia)
        {
          
            IEnumerable<SubFamiliaEntity> configuraciones = await _commonRepository.ListarSubFamilia(idlinea, idfamilia);
            ResponseModel<IEnumerable<SubFamiliaEntity>> resultadoConfiguraciones = new ResponseModel<IEnumerable<SubFamiliaEntity>>(true, Constante.MESSAGE_SUCCESS, configuraciones);

            return resultadoConfiguraciones;
        }

        public async Task<ResponseModel<IEnumerable<MarcaEntity>>> ListarMarca()
        {

            IEnumerable<MarcaEntity> configuraciones = await _commonRepository.ListarMarca();
            ResponseModel<IEnumerable<MarcaEntity>> resultadoConfiguraciones = new ResponseModel<IEnumerable<MarcaEntity>>(true, Constante.MESSAGE_SUCCESS, configuraciones);

            return resultadoConfiguraciones;
        }

        public async Task<ResponseModel<object>> RegistrarMaestroItem(DatosRequestMaestroItemModel dato, int idUsuario)
        {
            FormatoResponseRegistrarMaestroItem response = new FormatoResponseRegistrarMaestroItem();
            response = await _commonRepository.RegistrarMaestroItem(dato, idUsuario);
            return new ResponseModel<object>(true, Constante.MESSAGE_SUCCESS, new { response });
        }


        public async Task<PaginacionModel<FormatoListarMaestroItemModel>> ListarMaestroItem(DatosListarMaestroItemPaginador datos)
        {
            (List<FormatoListarMaestroItemModel> lista, int totalRegistros) = await _commonRepository.ListarMaestroItem(datos);

            PaginacionModel<FormatoListarMaestroItemModel> response = new PaginacionModel<FormatoListarMaestroItemModel>(lista, datos.Pagina, datos.RegistrosPorPagina, totalRegistros);

            return response;
        }

        public async Task<ResponseModel<IEnumerable<MaestroAlmacenEntity>>> ListarMaestroAlmacen()
        {

            IEnumerable<MaestroAlmacenEntity> MaestroAlmacen = await _commonRepository.ListarMaestroAlmacen();
            ResponseModel<IEnumerable<MaestroAlmacenEntity>> resultadoMaestroAlmacen = new ResponseModel<IEnumerable<MaestroAlmacenEntity>>(true, Constante.MESSAGE_SUCCESS, MaestroAlmacen);

            return resultadoMaestroAlmacen;
        }

        public async Task<ResponseModel<bool>> ValidacionPermisoAccesso(string Permiso, int idUsuario)
        {

            bool Acceso = await _commonRepository.ValidacionPermisoAccesso(Permiso, idUsuario);
            ResponseModel<bool> response = new ResponseModel<bool>(true, Constante.MESSAGE_SUCCESS, Acceso);

            return response;
        }

        public async Task<ResponseModel<IEnumerable<GrupoEntity>>> ListarGrupo()
        {
            IEnumerable<GrupoEntity> grupo = await _commonRepository.ListarGrupo();
            ResponseModel<IEnumerable<GrupoEntity>> resultado = new ResponseModel<IEnumerable<GrupoEntity>>(true, Constante.MESSAGE_SUCCESS, grupo);

            return resultado;
        }

        public async Task<ResponseModel<IEnumerable<TablaEntity>>> ListarTabla(string Grupo)
        {
            IEnumerable<TablaEntity> tabla = await _commonRepository.ListarTabla(Grupo);
            ResponseModel<IEnumerable<TablaEntity>> resultado = new ResponseModel<IEnumerable<TablaEntity>>(true, Constante.MESSAGE_SUCCESS, tabla);

            return resultado;
        }

        public async Task<ResponseModel<IEnumerable<MarcaProtocoloEntity>>> ListarMarcaProtocolo(string Grupo,string Campo)
        {
            IEnumerable<MarcaProtocoloEntity> response = await _commonRepository.ListarMarcaProtocolo(Grupo,Campo);
            ResponseModel<IEnumerable<MarcaProtocoloEntity>> resultado = new ResponseModel<IEnumerable<MarcaProtocoloEntity>>(true, Constante.MESSAGE_SUCCESS, response);

            return resultado;
        }

        public async Task<ResponseModel<IEnumerable<MetodologiaEntity>>> ListarMetodologiaProtocolo()
        {
            IEnumerable<MetodologiaEntity> response = await _commonRepository.ListarMetodologiaProtocolo();
            ResponseModel<IEnumerable<MetodologiaEntity>> resultadoMaestroAlmacen = new ResponseModel<IEnumerable<MetodologiaEntity>>(true, Constante.MESSAGE_SUCCESS, response);

            return resultadoMaestroAlmacen;
        }

        public async Task<ResponseModel<IEnumerable<AgrupadorHebrasEntity>>> ListarAgrupadoHebra()
        {
            IEnumerable<AgrupadorHebrasEntity> response = await _commonRepository.ListarAgrupadoHebra();
            ResponseModel<IEnumerable<AgrupadorHebrasEntity>> resultadoAgrupadoHebras = new ResponseModel<IEnumerable<AgrupadorHebrasEntity>>(true, Constante.MESSAGE_SUCCESS, response);

            return resultadoAgrupadoHebras;
        }

        public async Task<ResponseModel<IEnumerable<CalibrePruebaEntity>>> ListarCalibrePrueba()
        {
            IEnumerable<CalibrePruebaEntity> response = await _commonRepository.ListarCalibrePrueba();
            ResponseModel<IEnumerable<CalibrePruebaEntity>> resultadoAgrupadoHebras = new ResponseModel<IEnumerable<CalibrePruebaEntity>>(true, Constante.MESSAGE_SUCCESS, response);

            return resultadoAgrupadoHebras;
        }
        public async Task<ResponseModel<IEnumerable<TipoDocumentoSsomaEntity>>> TipoDocumentoSsoma()
        {

            IEnumerable<TipoDocumentoSsomaEntity> tipoDocumentoSsomas = await _commonRepository.TipoDocumentoSsoma();
            ResponseModel<IEnumerable<TipoDocumentoSsomaEntity>> resultadotipoDocumentoSsomas = new ResponseModel<IEnumerable<TipoDocumentoSsomaEntity>>(true, Constante.MESSAGE_SUCCESS, tipoDocumentoSsomas);

            return resultadotipoDocumentoSsomas;
        }

        public async Task<ResponseModel<IEnumerable<UbicacionSsomaEntity>>> UbicacionSsoma()
        {

            IEnumerable<UbicacionSsomaEntity> ubicacionSsomas = await _commonRepository.UbicacionSsoma();
            ResponseModel<IEnumerable<UbicacionSsomaEntity>> resultadoubicacionSsomas = new ResponseModel<IEnumerable<UbicacionSsomaEntity>>(true, Constante.MESSAGE_SUCCESS, ubicacionSsomas);

            return resultadoubicacionSsomas;
        }

        public async Task<ResponseModel<IEnumerable<ProteccionEntitySsoma>>> ProteccionSsoma()
        {

            IEnumerable<ProteccionEntitySsoma> ProteccionSsomas = await _commonRepository.ProteccionSsoma();
            ResponseModel<IEnumerable<ProteccionEntitySsoma>> resultadoProteccionSsomas = new ResponseModel<IEnumerable<ProteccionEntitySsoma>>(true, Constante.MESSAGE_SUCCESS, ProteccionSsomas);

            return resultadoProteccionSsomas;
        }
        public async Task<ResponseModel<IEnumerable<EstadoEntitySsoma>>> EstadoSsoma()
        {
            IEnumerable<EstadoEntitySsoma> EstadoSsomas = await _commonRepository.EstadoSsoma();
            ResponseModel<IEnumerable<EstadoEntitySsoma>> resultadoEstadoSsomas = new ResponseModel<IEnumerable<EstadoEntitySsoma>>(true, Constante.MESSAGE_SUCCESS, EstadoSsomas);
            return resultadoEstadoSsomas;
        }

        public async Task<ResponseModel<IEnumerable<AlmacenamientoSsomaEntity>>> AlmacenamientoSsoma()
        {
            IEnumerable<AlmacenamientoSsomaEntity> AlmacenamientoSsomas = await _commonRepository.AlmacenamientoSsoma();
            ResponseModel<IEnumerable<AlmacenamientoSsomaEntity>> Almacenamientosomas = new ResponseModel<IEnumerable<AlmacenamientoSsomaEntity>>(true, Constante.MESSAGE_SUCCESS, AlmacenamientoSsomas);
            return Almacenamientosomas;
        }

        public async Task<ResponseModel<IEnumerable<ResponsableSsomaEntity>>> ResponsableSsoma()
        {
            IEnumerable<ResponsableSsomaEntity> ResponsableSsomas = await _commonRepository.ResponsableSsoma();
            ResponseModel<IEnumerable<ResponsableSsomaEntity>> Responsablesomas = new ResponseModel<IEnumerable<ResponsableSsomaEntity>>(true, Constante.MESSAGE_SUCCESS, ResponsableSsomas);
            return Responsablesomas;
        }

        public async Task<ResponseModel<DatosClienteDTO>> ObtenerDatosCliente(int codigoCliente)
        {
            DatosClienteDTO datosCliente = await _commonRepository.ObtenerDatosCliente(codigoCliente);
            return new ResponseModel<DatosClienteDTO>(true, Constante.MESSAGE_SUCCESS, datosCliente); ;
        }

    }
}
