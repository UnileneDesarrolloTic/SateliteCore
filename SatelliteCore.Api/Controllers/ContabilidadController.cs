
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Services.Contracts;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;


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

        [HttpGet("ListarDetraccionContabilidad")]
        public async Task<ActionResult> ListarDetraccion()
        {

      
            IEnumerable<DetraccionesEntity> response = await _ContabilidadService.ListarDetraccion();

            return Ok(response);
        }

        [HttpGet("ProcesarDetraccionContabilidad")]
        public ActionResult ProcesarDetraccionContabilidad(string urlarchivo)
        {
            int response =  _ContabilidadService.ProcesarDetraccionContabilidad(urlarchivo);
            return Ok(response);
        }

       

        [HttpPost("GernerarBlogNotasDetraccion")]
        public ActionResult GenerarBlogNotasDetraccion(FormatoProcesoDetracciones dato)
        {
            string reporte =  _ContabilidadService.GenerarBlogNotasDetraccion(dato);
            return Ok(reporte);
        }



    }
}