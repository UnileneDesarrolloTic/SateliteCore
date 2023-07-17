using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SatelliteCore.Api.CrossCutting.Helpers;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Request.Dispensacion;
using SatelliteCore.Api.Models.Request.GestionGuias;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.Dispensacion;
using SatelliteCore.Api.Models.Response.Logistica;
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
    public class DispensacionController : ControllerBase
    {

        private readonly IDispensacionServices _dispensacionServices;

        public DispensacionController(IDispensacionServices dispensacionServices)
        {
            _dispensacionServices = dispensacionServices;
        }


        [HttpPost("ObtenerOrdenFabricacion")]
        public async Task<IActionResult> ObtenerOrdenFabricacion(DatosFormatoFiltroOrdenFabricacion dato)
        {
            ResponseModel<IEnumerable<DatosFormatoObtenerOrdenFabricacion>> listado = await _dispensacionServices.ObtenerOrdenFabricacion(dato);
            return Ok(listado);
        }

        [HttpGet("RecetasOrdenFabricacion")]
        public async Task<IActionResult> RecetasOrdenFabricacion(string ordenFabricacion)
        {
            IEnumerable<DatosFormatoListadoMateriaPrimaDispensacion> listado = await _dispensacionServices.RecetasOrdenFabricacion(ordenFabricacion);
            return Ok(listado);
        }

        [HttpPost("RegistrarDispensacionMP")]
        public async Task<IActionResult> RegistrarDispensacionMP(DatosFormatoDispensacionMateriaPrima dato)
        {
            string usuario = Shared.ObtenerUsuarioSpring(HttpContext.User.Identity);
            ResponseModel<string> respuesta = await _dispensacionServices.RegistrarDispensacionMP(dato, usuario);
            return Ok(respuesta);
        }

    }
}
