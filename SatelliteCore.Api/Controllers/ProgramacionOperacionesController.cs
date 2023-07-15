using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SatelliteCore.Api.CrossCutting.Helpers;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Request.GestionGuias;
using SatelliteCore.Api.Models.Request.ProgramacionOperaciones;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.Logistica;
using SatelliteCore.Api.Models.Response.ProgramacionOperaciones;
using SatelliteCore.Api.Services.Contracts;
using System.Collections.Generic;
using System.Threading.Tasks;
using SystemsIntegration.Api.Models.Exceptions;

namespace SatelliteCore.Api.Controllers
{
    [Authorize]
    [ApiController]
    [Route("api/[controller]")]
    [Produces("application/json")]
    public class ProgramacionOperacionesController : ControllerBase
    {
        private readonly IProgramacionOperacionesServices _programacionOperacionesServices;
        public ProgramacionOperacionesController(IProgramacionOperacionesServices programacionOperacionesServices)
        {
            _programacionOperacionesServices = programacionOperacionesServices;
        }

        [HttpGet("ObtenerAgrupadores")]
        public async Task<IActionResult> ObtenerAgrupadores(string gerencia)
        {
            IEnumerable<DatosFormatoAgrupadores> listado = await _programacionOperacionesServices.ObtenerAgrupadores(gerencia);
            return Ok(listado);
        }

        [HttpPost("ObtenerProgramacionOrdenFabricacion")]
        public async Task<IActionResult> ObtenerProgramacionOrdenFabricacion(DatosFormatoProgramacionOperaciones dato)
        {
            ResponseModel<IEnumerable<DatosFormatoProgramacionOperacionesOrdenFabricacion>> listado = await _programacionOperacionesServices.ObtenerProgramacionOrdenFabricacion(dato);
            return Ok(listado);
        }

        [HttpPost("ActualizarFechaProgramada")]
        public async Task<IActionResult> ActualizarFechaProgramada(DatosFormatoRegistrarFechaProgramacion dato)
        {
            string usuario = Shared.ObtenerUsuarioSpring(HttpContext.User.Identity);
            ResponseModel<string> respuesta = await _programacionOperacionesServices.ActualizarFechaProgramada(dato, usuario);
            return Ok(respuesta);
        }

        [HttpPost("RegistrarDivisionProgramacion")]
        public async Task<IActionResult> RegistrarDivisionProgramacion(DatosFormatoDividirRegistroProgramacion dato)
        {
            string usuario = Shared.ObtenerUsuarioSpring(HttpContext.User.Identity);
            ResponseModel<string> respuesta  = await _programacionOperacionesServices.RegistrarDivisionProgramacion(dato, usuario);
            return Ok(respuesta);
        }

    }
}
