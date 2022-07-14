using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Services.Contracts;
using System;
using System.Collections.Generic;
using System.Security.Claims;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Controllers
{
    [Authorize]
    [ApiController]
    [Route("api/[controller]")]
    [Produces("application/json")]
    public class LogisticaController : ControllerBase
    {
        private readonly ILogisticaServices _LogisticaService;
        public LogisticaController(ILogisticaServices LogisticaService)
        {
            _LogisticaService = LogisticaService;
        }

        [HttpPost("RegistrarMaestroItem")]
        public ActionResult RegistrarMaestroItem(DatosRequestMaestroItemModel dato)
        {
            try
            {
                _LogisticaService.RegistrarMaestroItem(dato);

                return Ok(dato);
            }
            catch (Exception ex)
            {
                ResponseModel<string> response
                        = new ResponseModel<string>(true, "No se completó ", ex.Message);
                return BadRequest(response);
            }


        }


    }
}
