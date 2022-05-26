
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
        public string GenerarBlogNotasDetraccion(FormatoProcesoDetracciones dato)
        {
            string valortext = "P20197705249UNILENE SAC                        " + dato.periodo + dato.totalimporte.ToString("D15") + "\n";

            foreach (var detraccion in dato.proceso)
            {
                var numero = Int32.Parse(detraccion.Numero);
                var servicio = Int32.Parse(detraccion.BienServicio);
                var cuenta = "00002007770";
                var operacion = Int32.Parse(detraccion.Tipo);
                var tipo = Int32.Parse(detraccion.Tipo);

                valortext += detraccion.TipoDocumento + "" + detraccion.Ruc + "                                   " + servicio.ToString("D12") + cuenta + detraccion.Importe.ToString("D13") + operacion.ToString("D4") + detraccion.Periodo + tipo.ToString("D2") + detraccion.Serie + numero.ToString("D8") + "\n";
            }
            var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(valortext);

            return System.Convert.ToBase64String(plainTextBytes);
        }



    }
}