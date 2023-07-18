using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.CrossCutting.Helpers;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.Common;
using SatelliteCore.Api.Services.Contracts;
using System.Collections.Generic;
using System.Security.Claims;
using System.Threading.Tasks;
using SystemsIntegration.Api.Models.Exceptions;

namespace SatelliteCore.Api.Controllers
{
    [Authorize]
    [ApiController]
    [Route("api/[controller]")]
    [Produces("application/json")]
    public class CommonController : ControllerBase
    {
        private readonly ICommonServices _commonService;
        public CommonController(ICommonServices commonService)
        {
            _commonService = commonService;
        }

        [HttpGet("ListarTipoDocumentoIdentidad")]
        public async Task<ActionResult> ListarTipoDocumentoIdentidad()
        {
            IEnumerable<TipoDocumentoIdentidadEntity> lista = await _commonService.ListarTipoDocumentoIndentidad();
            return Ok(lista);
        }

        [HttpGet("ListarPaises")]
        public async Task<ActionResult> ListarPaises()
        {
            IEnumerable<PaisEntity> lista = await _commonService.ListarPaises();

            return Ok(lista);
        }


        [HttpGet("ObtenerMenuUsuarioSesion")]
        public async Task<ActionResult> ObtenerMenuUsuarioSesion()
        {
            var claims = HttpContext.User.Identity as ClaimsIdentity;
            int idUsuario = int.Parse(claims.FindFirst(ClaimTypes.NameIdentifier).Value);
            List<MenuxUsuarioModel> lista = await _commonService.ListarMenuxUsuario(idUsuario);

            return Ok(lista);
        }

        [HttpGet("listarRoles")]
        public async Task<ActionResult> ListarRoles(string estado)
        {
            IEnumerable<RolEntity> listaUsuario = await _commonService.ListarRoles(estado);
            return Ok(listaUsuario);
        }


        [HttpGet("ListarFamiliaMP")]
        public async Task<ActionResult> ListarFamiliaMP(string tipo)
        {
            List<FamiliaMP> listarFamiliaMP = await _commonService.ListarFamiliaMP(tipo);
            return Ok(listarFamiliaMP);
        }

        [HttpGet("ObtenerConfiguracionesSistema")]
        public async Task<IActionResult> ObtenerConfiguracionesSistema(int idConfiguracion, string grupo)
        {
            ResponseModel<IEnumerable<ConfiguracionEntity>> configuraciones = await _commonService.ObtenerConfiguracionesSistema(idConfiguracion, grupo);
            return Ok(configuraciones);
        }


        [HttpGet("ListarAgrupador")]
        public async Task<IActionResult> ListarAgrupador()
        {
            ResponseModel<IEnumerable<AgrupadorEntity>> configuraciones = await _commonService.ListarAgrupador();
            return Ok(configuraciones);
        }

        [HttpGet("ListarSubAgrupador")]
        public async Task<IActionResult> ListarSubAgrupador(string idAgrupador)
        {
            ResponseModel<IEnumerable<SubAgrupadorEntity>> configuraciones = await _commonService.ListarSubAgrupador(idAgrupador);
            return Ok(configuraciones);
        }

        [HttpGet("ListarLinea")]
        public async Task<IActionResult> ListarLinea()
        {
            ResponseModel<IEnumerable<LineaEntity>> configuraciones = await _commonService.ListarLinea();
            return Ok(configuraciones);
        }

        [HttpGet("ListarFamilia")]
        public async Task<IActionResult> ListarFamilia(string idLinea)
        {
            ResponseModel<IEnumerable<FamiliaMaestroItemsModel>> configuraciones = await _commonService.ListarFamilia(idLinea);
            return Ok(configuraciones);
        }

        [HttpGet("ListarFamiliaGeneral")]
        public async Task<IActionResult> ListarFamiliaGeneral(string idLinea)
        {
            ResponseModel<IEnumerable<FamiliaMaestroItemsModel>> configuraciones = await _commonService.ListarFamiliaGeneral(idLinea);
            return Ok(configuraciones);
        }

        [HttpGet("ListarSubFamilia")]
        public async Task<IActionResult> ListarSubFamilia(string idlinea, string idfamilia)
        {
            ResponseModel<IEnumerable<SubFamiliaEntity>> configuraciones = await _commonService.ListarSubFamilia(idlinea,idfamilia);
            return Ok(configuraciones);
        }

        [HttpGet("ListarMarca")]
        public async Task<IActionResult> ListarMarca()
        {
            ResponseModel<IEnumerable<MarcaEntity>> configuraciones = await _commonService.ListarMarca();
            return Ok(configuraciones);
        }

        [HttpPost("RegistrarMaestroItem")]
        public async Task<ActionResult> RegistrarMaestroItem(DatosRequestMaestroItemModel dato)
        {
            int idUsuario = Shared.ObtenerUsuarioSesion(HttpContext.User.Identity);
            ResponseModel<object> cantidadPruebas = await _commonService.RegistrarMaestroItem(dato, idUsuario);
            return Ok(cantidadPruebas);
        }


        [HttpPost("ListarMaestroItem")]
        public async Task<IActionResult> ListarMaestroItem(DatosListarMaestroItemPaginador datos)
        {
            PaginacionModel<FormatoListarMaestroItemModel> response = await _commonService.ListarMaestroItem(datos);
            return Ok(response);
        }

        [HttpGet("ListarMaestroAlmacen")]
        public async Task<IActionResult> ListarMaestroAlmacen()
        {
            ResponseModel<IEnumerable<MaestroAlmacenEntity>> MaestroAlmacen = await _commonService.ListarMaestroAlmacen();
            return Ok(MaestroAlmacen);
        }

        [HttpGet("ValidacionPermisoAccesso")]
        public async Task<IActionResult> ValidacionPermisoAccesso(string Permiso)
        {   
            if(string.IsNullOrEmpty(Permiso))
                throw new ValidationModelException("verificar el parametro enviado");

            int idUsuario = Shared.ObtenerUsuarioSesion(HttpContext.User.Identity);
            ResponseModel<bool> MaestroAlmacen = await _commonService.ValidacionPermisoAccesso(Permiso, idUsuario);
            return Ok(MaestroAlmacen);
        }


        [HttpGet("ListarGrupo")]
        public async Task<IActionResult> ListarGrupo()
        {
            ResponseModel<IEnumerable<GrupoEntity>> response = await _commonService.ListarGrupo();
            return Ok(response);
        }

        [HttpGet("ListarTabla")]
        public async Task<IActionResult> ListarTabla(string Grupo)
        {
            ResponseModel<IEnumerable<TablaEntity>> response = await _commonService.ListarTabla(Grupo);
            return Ok(response);
        }


        [HttpGet("ListarMarcaProtocolo")]
        public async Task<IActionResult> ListarMarcaProtocolo(string Grupo, string Campo)
        {
            ResponseModel<IEnumerable<MarcaProtocoloEntity>> response = await _commonService.ListarMarcaProtocolo(Grupo, Campo);
            return Ok(response);
        }

        [HttpGet("ListarMetodologiaProtocolo")]
        public async Task<IActionResult> ListarMetodologiaProtocolo()
        {
            ResponseModel<IEnumerable<MetodologiaEntity>> response = await _commonService.ListarMetodologiaProtocolo();
            return Ok(response);
        }

        [HttpGet("ListarAgrupadoHebra")]
        public async Task<IActionResult> ListarAgrupadoHebra()
        {
            ResponseModel<IEnumerable<AgrupadorHebrasEntity>> response = await _commonService.ListarAgrupadoHebra();
            return Ok(response);
        }

        [HttpGet("ListarCalibrePrueba")]
        public async Task<IActionResult> ListarCalibrePrueba()
        {
            ResponseModel<IEnumerable<CalibrePruebaEntity>> response = await _commonService.ListarCalibrePrueba();
            return Ok(response);
        }


        [HttpGet("TipoDocumentoSsoma")]
        public async Task<IActionResult> TipoDocumentoSsoma()
        {
            ResponseModel<IEnumerable<TipoDocumentoSsomaEntity>> TipoDocumentoSsoma = await _commonService.TipoDocumentoSsoma();
            return Ok(TipoDocumentoSsoma);
        }

        [HttpGet("UbicacionSsoma")]
        public async Task<IActionResult> UbicacionSsoma()
        {
            ResponseModel<IEnumerable<UbicacionSsomaEntity>> UbicacionSsoma = await _commonService.UbicacionSsoma();
            return Ok(UbicacionSsoma);
        }

        [HttpGet("ProteccionSsoma")]
        public async Task<IActionResult> ProteccionSsoma()
        {
            ResponseModel<IEnumerable<ProteccionEntitySsoma>> ProteccionSsoma = await _commonService.ProteccionSsoma();
            return Ok(ProteccionSsoma);
        }

        [HttpGet("EstadoSsoma")]
        public async Task<IActionResult> EstadoSsoma()
        {
            ResponseModel<IEnumerable<EstadoEntitySsoma>> EstadoSsoma = await _commonService.EstadoSsoma();
            return Ok(EstadoSsoma);
        }

        [HttpGet("AlmacenamientoSsoma")]
        public async Task<IActionResult> AlmacenamientoSsoma()
        {
            ResponseModel<IEnumerable<AlmacenamientoSsomaEntity>> AlmacenamientoSsoma = await _commonService.AlmacenamientoSsoma();
            return Ok(AlmacenamientoSsoma);
        }

        [HttpGet("ResponsableSsoma")]
        public async Task<IActionResult> ResponsableSsoma()
        {
            ResponseModel<IEnumerable<ResponsableSsomaEntity>> AlmacenamientoSsoma = await _commonService.ResponsableSsoma();
            return Ok(AlmacenamientoSsoma);
        }

        [HttpGet("DatosCliente")]
        public async Task<IActionResult> ObtenerDatosCliente(int codigoCliente)
        {
            ResponseModel<DatosClienteDTO> cliente = await _commonService.ObtenerDatosCliente(codigoCliente);
            return Ok(cliente);
        }

        [HttpGet("Transportista")]
        public async Task<IActionResult> Transportista()
        {   
            ResponseModel<IEnumerable<TransportistaEntity>> transportista = await _commonService.Transportista();
            return Ok(transportista);
        }

        [HttpGet("ClasificacionArea")]
        public async Task<IActionResult> ClasificacionArea()
        {
            ResponseModel<IEnumerable<ClasificacionAreaEntity>> clasificacionArea = await _commonService.ClasificacionArea();
            return Ok(clasificacionArea);
        }

        [HttpGet("InformacionItem")]
        public async Task<IActionResult> InformacionItem(string item)
        {
            ResponseModel<DatosFormatoInformacionItem> clasificacionArea = await _commonService.InformacionItem(item);
            return Ok(clasificacionArea);
        }
    }
}
