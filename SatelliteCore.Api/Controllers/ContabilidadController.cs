using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SatelliteCore.Api.CrossCutting.Helpers;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Services.Contracts;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using SystemsIntegration.Api.Models.Exceptions;

namespace SatelliteCore.Api.Controllers
{
    [Authorize]
    [ApiController]
    [Route("api/[controller]")]
    [Produces("application/json")]
    public class ContabilidadController : ControllerBase
    {
        private readonly IContabilidadService _ContabilidadService;
        
        public ContabilidadController(IContabilidadService ContabilidadService)
        {
            _ContabilidadService = ContabilidadService;
           
        }

        [HttpPost("ListarDetraccionContabilidad")]
        public async Task<ActionResult> ListarDetraccion(DatosListarDetraccionPaginado datos)
        {
            if (!ModelState.IsValid)
            {
                ResponseModel<string> responseError = new ResponseModel<string>(false, "Los datos enviado no son válidos.", "");

                return BadRequest(responseError);
            }

            PaginacionModel<DetraccionesEntity> response = await _ContabilidadService.ListarDetraccion(datos);
            return Ok(response);
        }
            


    }
}
