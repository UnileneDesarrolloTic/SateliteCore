using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SatelliteCore.Api.CrossCutting.Helpers;
using SatelliteCore.Api.Models.Dto.AnalisisAgujas;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Services.Contracts;
using System;
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
    public class LogisticaController : ControllerBase
    {
        private readonly ILogisticaServices _logisticaServices;

        public LogisticaController(ILogisticaServices logisticaServices)
        {
            _logisticaServices = logisticaServices;
        }

        [HttpGet("ObtenerNumeroGuias")]
        public async Task<IActionResult> ObtenerNumeroGuias(string NumeroGuia)
        {
            if (NumeroGuia == "")
                throw new ValidationModelException("Debe Ingresar el numero de la guia");

            IEnumerable<DatosFormatoPlanOrdenServicosD> listaAnalisis = await _logisticaServices.ObtenerNumeroGuias(NumeroGuia);

            return Ok(listaAnalisis);
        }

        [HttpPost("RegistrarRetornoGuia")]
        public async Task<IActionResult> RegistrarRetornoGuia(List<DatosFormatoRetornoGuiaRequest> dato )
        {
            if (!ModelState.IsValid)
                throw new ValidationModelException("Los datos de prueba de flexion no son válidos !!");

            ResponseModel<string> response = await _logisticaServices.RegistrarRetornoGuia(dato);

            return Ok(response);
        }

    }
}
