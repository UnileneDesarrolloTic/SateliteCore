using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.CrossCutting.Helpers;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Services.Contracts;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Controllers
{
    [Authorize]
    [ApiController]
    [Route("api/[controller]")]
    [Produces("application/json")]
    public class UsuarioController : ControllerBase
    {
        private readonly IUsuarioService _usuarioService;
        public UsuarioController(IUsuarioService usuarioService)
        {
            _usuarioService = usuarioService;
        }


        [HttpPost("ObtenerUsuario")]
        public async Task<ActionResult> ObtenerUsuario(DatoUsuarioBasico user)
        {
            if (!ModelState.IsValid)
            {
                //IEnumerable<string> errorList = ModelState.Values.SelectMany(m => m.Errors).Select(e => e.ErrorMessage).ToList();
                ResponseModel<string> responseError =
                        new ResponseModel<string>(false, Constante.MODEL_VALIDATION_FAILED, "Los datos enviados no son válidos");

                return BadRequest(responseError);
            }

            UsuarioEntity usuario = await _usuarioService.ObtenerUsuario(user);

            if (usuario == null)
            {
                ResponseModel<string> response = new ResponseModel<string>(false, "No se pudo encontrar al usuario", "");
                return Ok(response);
            }

            ResponseModel<UsuarioEntity> responseSuccesss =
                new ResponseModel<UsuarioEntity>(true, Constante.MESSAGE_SUCCESS, usuario);

            return Ok(responseSuccesss);
        }


        [HttpPost("ListarUsuarios")]
        public async Task<ActionResult> ListarUsuario(DatosListarUsuarioPaginado datos)
        {
            if (!ModelState.IsValid)
            {
                ResponseModel<string> responseError =
                        new ResponseModel<string>(false, Constante.MODEL_VALIDATION_FAILED, "");

                return BadRequest(responseError);
            }

            PaginacionModel<UsuarioEntity> response = await _usuarioService.ListarUsuarios(datos);

            return Ok(response);

        }

        [HttpPost("CambioClave")]
        public async Task<ActionResult> CambiarClaveUsuario(ActualizarClaveModel datos)
        {
            if (!ModelState.IsValid)
            {
                ResponseModel<string> responseError =
                        new ResponseModel<string>(false, Constante.MODEL_VALIDATION_FAILED, "");

                return BadRequest(responseError);
            }

            int result = await _usuarioService.CambiarClave(datos);

            ResponseModel<int> response
                    = new ResponseModel<int>(true, "La clave se ha cambiado correcamente", result);

            return Ok(response);

        }


        [HttpGet("ListarAsignacionPersonal")]
        public async Task<ActionResult> ListarAsignacionPersonal()
        {
            DatosFormatoAsignacionPersonalLaboralModel response = await _usuarioService.ListarAsignacionPersonal();
            return Ok(response);
        }

        [HttpGet("ListarAreaPersonaLaboral")]
        public async Task<ActionResult> ListarAreaPersonaLaboral()
        {
            List<AreaPersonalLaboralEntity> response = await _usuarioService.ListarAreaPersonaLaboral();
            return Ok(response);
        }

        [HttpPost("RegistrarPersonaMasiva")]
        public async Task<ActionResult> RegistrarPersonaLaboralMasiva(DatosFormatoAsignacionPersonalModel dato)
        {
            int idUsuario = Shared.ObtenerUsuarioSesion(HttpContext.User.Identity);

            ResponseModel<string> response = await _usuarioService.RegistrarPersonaLaboralMasiva(dato, idUsuario);
            return Ok(response);
        }

        [HttpGet("FiltrarAreaPersona")]
        public async Task<ActionResult> FiltrarAreaPersona(int idArea, string NombrePersona)
        {
            IEnumerable<DatosFormatoFiltrarTrabajadorAreaModel> response = await _usuarioService.FiltrarAreaPersona(idArea, NombrePersona);
            return Ok(response);
        }

        [HttpGet("LiberalPersona")]
        public async Task<ActionResult> LiberalPersona(int IdAsignacion)
        {
            ResponseModel<string> response = await _usuarioService.LiberalPersona(IdAsignacion);
            return Ok(response);
        }


        [HttpGet("ExportarExcelPersonaAsignacion")]
        public async Task<ActionResult> ExportarExcelPersonaAsignacion(string FechaInicio , string FechaFinal)
        {   
            ResponseModel<string> response = await _usuarioService.ExportarExcelPersonaAsignacion(FechaInicio, FechaFinal);
            return Ok(response);
        }


        [HttpGet("RegistrarEditarArea")]
        public async Task<ActionResult> RegistrarEditarArea(int IdArea, string Descripcion)
        {
            ResponseModel<AreaPersonalLaboralEntity> response = await _usuarioService.RegistrarEditarArea(IdArea, Descripcion);
            return Ok(response);
        }

        [HttpGet("EliminarAreaProduccion")]
        public async Task<ActionResult> EliminarAreaProduccion(int IdArea)
        {
            ResponseModel<string> response = await _usuarioService.EliminarAreaProduccion(IdArea);
            return Ok(response);
        }

        [HttpGet("ListarPersonaTecnico")]
        public async Task<ActionResult> ListarPersonaTecnico()
        {
            IEnumerable<DatosFormatoListarPersonaTecnica> response = await _usuarioService.ListarPersonaTecnico();
            return Ok(response);
        }


    }
}
