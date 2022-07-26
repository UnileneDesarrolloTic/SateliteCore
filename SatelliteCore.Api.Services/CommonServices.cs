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

        public async Task<ResponseModel<object>> RegistrarMaestroItem(DatosRequestMaestroItemModel dato)
        {
            FormatoResponseRegistrarMaestroItem response = new FormatoResponseRegistrarMaestroItem();
            response = await _commonRepository.RegistrarMaestroItem(dato);
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



    }
}
