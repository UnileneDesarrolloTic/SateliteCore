using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Services.Contracts;
using System.Collections.Generic;
using System.Security.Claims;
using System.Threading.Tasks;

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

        [HttpGet("ListarSubFamilia")]
        public async Task<IActionResult> ListarSubFamilia(string idlinea, string idfamilia)
        {
            ResponseModel<IEnumerable<SubFamiliaEntity>> configuraciones = await _commonService.ListarSubFamilia(idlinea,idfamilia);
            return Ok(configuraciones);
        }




    }
}
