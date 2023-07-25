using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SatelliteCore.Api.CrossCutting.Helpers;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.GestioEquipoEngaste;
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
    public class GestionEquipoEngasteController : ControllerBase
    {
        private readonly IGestionEquipoEngasteServices _gestionEquipoEngasteServices;

        public GestionEquipoEngasteController(IGestionEquipoEngasteServices gestionEquipoEngasteServices)
        {
            _gestionEquipoEngasteServices = gestionEquipoEngasteServices;
        }

        [HttpGet("ObtenerEmpleado")]
        public async Task<IActionResult> ObtenerEmpleado()
        {
            IEnumerable<DatosFormatoEmpleado> listado = await _gestionEquipoEngasteServices.ObtenerEmpleado();
            return Ok(listado);
        }

        [HttpGet("ObtenerListadoDados")]
        public async Task<IActionResult> ObtenerListadoDados()
        {
            IEnumerable<DatosFormatoListadoDadoEngaste> listado = await _gestionEquipoEngasteServices.ObtenerListadoDados();
            return Ok(listado);
        }

        [HttpPost("ListarEquipoEngaste")]
        public async Task<IActionResult> ListarEquipoEngaste(DatosFormularioFiltroEquipo dato)
        {
            IEnumerable<DatosFormatoListarEquipoEngaste> listado = await _gestionEquipoEngasteServices.ListarEquipoEngaste(dato);
            return Ok(listado);
        }

        [HttpGet("ObtenerInformacionEquipo")]
        public async Task<IActionResult> ObtenerInformacionEquipo(string idEquipo)
        {
            DatosFormatoInformacionEquipoEngaste informacion = await _gestionEquipoEngasteServices.ObtenerInformacionEquipo(idEquipo);
            return Ok(informacion);
        }


        [HttpPost("RegistrarEquipoEngastado")]
        public async Task<IActionResult> RegistrarEquipoEngastado(DatosFormatoRegistroEquipoEngastado dato)
        {
            ResponseModel<string> respuesta = await _gestionEquipoEngasteServices.RegistrarEquipoEngastado(dato);
            return Ok(respuesta);
        }


    }
}
