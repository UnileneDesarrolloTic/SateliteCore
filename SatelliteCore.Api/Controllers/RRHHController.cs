using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SatelliteCore.Api.CrossCutting.Helpers;
using SatelliteCore.Api.Models.Config;
using SatelliteCore.Api.Models.Dto.RRHH;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Services.Contracts;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Controllers
{
    [Authorize]
    [ApiController]
    [Route("api/[controller]")]
    [Produces("application/json")]
    public class RRHHController : ControllerBase
    {
        private readonly IAppConfig _appConfig;
        private readonly IRRHHServices _rrhhServices;

        public RRHHController(IAppConfig appConfig, IRRHHServices rrhhServices)
        {
            _appConfig = appConfig;
            _rrhhServices = rrhhServices;
        }

        [HttpPost("GenerarReporteAsistencia")]
        public async Task<ActionResult> GenerarReporteAsistencia(DatosReporteRRHH datos)
        {
            try
            {
                string Reporte = "ReporteDiarioAsistencia&rs:Command=Render";
                string Formato = "&rs:Format=excel";
                string Parametros = "&Fecha=" + datos.FechaReporte.ToString().Substring(0, 10);


                var theURL = _appConfig.ReportRRHH + Reporte + Parametros + Formato;


                var httpClientHandler = new HttpClientHandler()
                {
                    UseDefaultCredentials = true
                };

                HttpClient webClient = new HttpClient(httpClientHandler);

                Byte[] result = await webClient.GetByteArrayAsync(theURL);
                string base64String = Convert.ToBase64String(result, 0, result.Length);
                ResponseModel<string> response
                        = new ResponseModel<string>(true, "El reporte se generó correctamente", base64String);
                return Ok(response);
            }
            catch (Exception ex)
            {
                ResponseModel<string> response
                        = new ResponseModel<string>(true, "El reporte no se generó", ex.Message);
                return BadRequest(response);
            }

        }

        [HttpGet("ListarAsistencia")]
        public async Task<IActionResult> ListarAsistencia(DateTime fecha)
        {
            int usuarioToken = Shared.ObtenerUsuarioSesion(HttpContext.User.Identity);
            ResponseModel<IEnumerable<ReporteAsistenciaDTO>> listarAsistencia = await _rrhhServices.ListarAsistencia(fecha, usuarioToken);
            return Ok(listarAsistencia);
        }

        [HttpPost("RegistrarHorasExtras")]
        public async Task<IActionResult> RegistrarHorasExtras(DatosEstructuraHorasExtraCabecera dato)
        {
            string usuario = Shared.ObtenerUsuarioSpring(HttpContext.User.Identity);

            ResponseModel<string> response = await _rrhhServices.RegistrarHorasExtras(dato, usuario);
            return Ok(response);
        }

        [HttpPost("ListarHoraExtrasPersona")]
        public async Task<IActionResult> ListarHoraExtrasPersona(DatosFormatoFiltraHoraExtras dato)
        {
            IEnumerable<DatosFormatoListarHorasExtrasPersona> listado = await _rrhhServices.ListarHoraExtrasPersona(dato);
            return Ok(listado);
        }

        [HttpGet("ProcesarHorasExtrasPlanilla")]
        public async Task<IActionResult> ProcesarHorasExtrasPlanilla(string periodo)
        {
            await _rrhhServices.ProcesarHorasExtrasPlanilla(periodo);
            return Ok();
        }

        [HttpGet("InformacionHoraExtras")]
        public async Task<IActionResult> BuscarInformacionHorasExtrasPersona(int Cabecera)
        {
            object listado = await _rrhhServices.BuscarInformacionHorasExtrasPersona(Cabecera);
            return Ok(listado);
        }

        [HttpGet("ReporteHorasExtrasGenerada")]
        public async Task<IActionResult> ReporteHorasExtrasGenerada(string periodo)
        {
            string reporte =  await _rrhhServices.ReporteHorasExtrasGeneradas(periodo);
            return Ok(reporte);
        }

        [HttpGet("FormatoAutorizacionSobretiempo")]
        public async Task<IActionResult> FormatoAutorizacionSobretiempo(int id)
        {
            string reporte = await _rrhhServices.FormatoAutorizacionSobretiempo(id);
            return Ok(reporte);
        }

        [HttpGet("RptAutorizacionSobretiempoPorPersona")]
        public async Task<IActionResult> AutorizacionSobretiempoPorPersona(string periodo)
        {
            ResponseModel<string> reporte = await _rrhhServices.AutorizacionSobretiempoPorPersona(periodo);
            return Ok(reporte);
        }


    }
}
