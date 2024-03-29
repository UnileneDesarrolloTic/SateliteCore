﻿using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Services.Contracts;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Controllers.Auth
{

    [ApiController]
    [Authorize]
    [Route("api/[controller]")]
    [Produces("application/json")]
    public class ValidacionAccesosController : ControllerBase
    {
        private readonly IValidacionesServices _validacionesServices;

        public ValidacionAccesosController(IValidacionesServices validacionesServices)
        {
            _validacionesServices = validacionesServices;
        }

        [HttpPost("ValidarAccesoRuta")]
        public async Task<IActionResult> AuthenticateUser(ValidacionRutaDataModel model)
        {

            if (!ModelState.IsValid)
            {
                var listaErrores = ModelState.Values.SelectMany(m => m.Errors).Select(e => e.ErrorMessage).ToList();
                ResponseModel<IEnumerable<string>> responseError =
                        new ResponseModel<IEnumerable<string>>(false, Constante.MODEL_VALIDATION_FAILED, listaErrores);
                return BadRequest(responseError);
            }

            var hola = ModelState["values"];

            (int codigo, string mensaje) = await _validacionesServices.ValidarAccesoRuta(model);

            ResponseModel<object> responseSuccesss = new ResponseModel<object>(true, Constante.MESSAGE_SUCCESS, new { codigo, mensaje });
            return Ok(responseSuccesss);
        }

        //[PermitCode("FR0013")]
        [HttpGet("validarToken")]
        public ActionResult ValidarToken()
        {
            return Ok("OK");
        }

    }
}
