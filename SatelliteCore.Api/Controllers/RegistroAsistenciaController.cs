using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Services.Contracts;
using SatelliteCore.Api.CrossCutting.Config;

namespace SatelliteCore.Api.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    [Produces("application/json")]
    public class RegistroAsistenciaController : ControllerBase
    {
        private readonly IRegistroAsistenciaServices _registroAsistenciaServices;
        public RegistroAsistenciaController(IRegistroAsistenciaServices registroAsistenciaServices)
        {
            _registroAsistenciaServices = registroAsistenciaServices;
        }

        [HttpGet("RegistraAsistencia")]
        public async Task<IActionResult> RegistraAsistencia(string numeroDocumento)
        {
            ResponseModel<string> response = await _registroAsistenciaServices.RegistraAsistencia(numeroDocumento);
            return Ok(response);
        }
    }
}
