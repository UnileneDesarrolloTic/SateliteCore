using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.CrossCutting.Helpers;
using SatelliteCore.Api.Models.Config;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Services.Contracts;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Security.Claims;
using System.Threading.Tasks;
using SystemsIntegration.Api.Models.Exceptions;

namespace SatelliteCore.Api.Controllers
{
    [Authorize]
    [ApiController]
    [Route("api/[controller]")]
    [Produces("application/json")]
    public class ControlCalidadController : ControllerBase
    {
        private readonly IControlCalidadServices _controlCalidadServices;
        private readonly IAppConfig _appConfig;
        public ControlCalidadController(IControlCalidadServices controlCalidadServices, IAppConfig appConfig)
        {
            _controlCalidadServices = controlCalidadServices;
            _appConfig = appConfig;
        }

        [HttpPost("ListarCertificados")]
        public async Task<ActionResult> ListarCertificados(DatosListarCertificadoPaginado datos)
        {
            try
            {
                if (!ModelState.IsValid)
                {
                    ResponseModel<string> responseError =
                            new ResponseModel<string>(false, Constante.MODEL_VALIDATION_FAILED, "");

                    return BadRequest(responseError);
                }

                (List<CertificadoEsterilizacionEntity> lista, int totalRegistros) = await _controlCalidadServices.ListarCertificados(datos);

                PaginacionModel<CertificadoEsterilizacionEntity> response
                        = new PaginacionModel<CertificadoEsterilizacionEntity>(lista, datos.Pagina, datos.RegistrosPorPagina, totalRegistros);
                return Ok(response);
            }
            catch (Exception ex)
            {

                return BadRequest(ex.Message);
            }

        }

        [HttpPost("RegistrarCertificado")]
        public ActionResult RegistrarCertificado(CertificadoEsterilizacionEntity certificado)
        {
            try
            {
                _controlCalidadServices.RegistrarCertificado(certificado);

                return Ok();
            }
            catch (Exception ex)
            {
                ResponseModel<string> response
                        = new ResponseModel<string>(true, "No se completó ", ex.Message);
                return BadRequest(response);
            }


        }

        [HttpPost("ListarLotes")]
        public async Task<ActionResult> ListarLotes(DatosLote datos)
        {
            if (!ModelState.IsValid)
            {
                ResponseModel<string> responseError =
                        new ResponseModel<string>(false, Constante.MODEL_VALIDATION_FAILED, "");

                return BadRequest(responseError);
            }

            (List<LoteEntity> lista, int totalRegistros) = await _controlCalidadServices.ListarLotes(datos);

            PaginacionModel<LoteEntity> response
                    = new PaginacionModel<LoteEntity>(lista, datos.Pagina, datos.RegistrosPorPagina, totalRegistros);

            return Ok(response);
        }

        [HttpPost("GenerarReporte")]
        public async Task<ActionResult> GenerarReporte(DatosReporte datos)
        {
            try
            {
                string ReporteEsterilizacion = "Certificado_Esterilizacion&rs:Command=Render";
                string Formato = "&rs:Format=PDF";
                string Parametros = "&Id=" + datos.Id;

                var theURL = _appConfig.ReportControlDeCalidad + ReporteEsterilizacion + Parametros + Formato;

                var httpClientHandler = new HttpClientHandler()
                {
                    UseDefaultCredentials = true
                };

                HttpClient webClient = new HttpClient(httpClientHandler);

                Byte[] result = await webClient.GetByteArrayAsync(theURL);
                string base64String = Convert.ToBase64String(result, 0, result.Length);
                ResponseModel<string> response = new ResponseModel<string>(true, "El reporte se generó correctamente", base64String);
                return Ok(response);
            }
            catch (Exception ex)
            {
                ResponseModel<string> response
                        = new ResponseModel<string>(true, "El reporte no se generó", ex.Message);
                return BadRequest(response);
            }

        }

        [HttpPost("RegistrarLote")]
        public async Task<ActionResult> RegistrarLote(LoteEntity lote)
        {
            int result = await _controlCalidadServices.RegistrarLote(lote);

            ResponseModel<int> response
                    = new ResponseModel<int>(true, "El lote se registró correctamente", result);

            return Ok(response);

        }

        [HttpPost("ListarCotizaciones")]
        public async Task<ActionResult> ListarCotizaciones(DatosListarCotizacionesPaginado datos)
        {
            if (!ModelState.IsValid)
            {
                ResponseModel<string> responseError =
                        new ResponseModel<string>(false, Constante.MODEL_VALIDATION_FAILED, "");

                return BadRequest(responseError);
            }

            (List<CotizacionEntity> lista, int totalRegistros) = await _controlCalidadServices.ListarCotizaciones(datos);

            PaginacionModel<CotizacionEntity> response
                    = new PaginacionModel<CotizacionEntity>(lista, datos.Pagina, datos.RegistrosPorPagina, totalRegistros);

            return Ok(response);
        }

        [HttpGet("ObtenerInformacionLote")]
        public async Task<ActionResult> ObtenerInformacionLote(string NumeroLote)
        {
            if (NumeroLote == "")
            {
                throw new ValidationModelException("Debe Ingresar el Numero de Lote");
            }

            IEnumerable<FormatoEstructuraObtenerOrdenFabricacion> response =  await _controlCalidadServices.ObtenerInformacionLote(NumeroLote);
            return Ok(response);
        }


        [HttpGet("ListarTransaccionItem")]
        public async Task<ActionResult> ListarTransaccionItem(string NumeroLote, string codAlmacen)
        {
            if (NumeroLote == "")
            {
                ResponseModel<string> responseError = new ResponseModel<string>(false, Constante.MODEL_VALIDATION_FAILED, "");
                return BadRequest(responseError);
            }

            IEnumerable<DatosFormatoListarTransaccion> response = await _controlCalidadServices.ListarTransaccionItem(NumeroLote, codAlmacen);
            return Ok(response);
        }


        [HttpPost("RegistrarLoteNumeroCaja")]
        public async Task<ActionResult> RegistrarLoteNumeroCaja(DatosFormatoOrdenFabricacionRequest dato)
        {
            int idUsuario = Shared.ObtenerUsuarioSesion(HttpContext.User.Identity);
            ResponseModel<string> response = await _controlCalidadServices.RegistrarLoteNumeroCaja(dato, idUsuario);
            return Ok(response);
        }

        [HttpGet("ListarKardexInternoNumeroLote")]
        public async Task<ActionResult> ListarKardexInternoNumeroLote(string NumeroLote)
        {
            IEnumerable < DatosFormatoKardexInternoGCM > response= await _controlCalidadServices.ListarKardexInternoNumeroLote(NumeroLote);
            return Ok(response);
        }

        [HttpPost("RegistrarKardexInternoGCM")]
        public async Task<ActionResult> RegistrarKardexInternoGCM(DatosFormatoRegistrarKardexInternoGCM dato)
        {
            int idUsuario = Shared.ObtenerUsuarioSesion(HttpContext.User.Identity);
            ResponseModel<string> response = await _controlCalidadServices.RegistrarKardexInternoGCM(dato, idUsuario);
            return Ok(response);
        }

        [HttpGet("ActualizarKardexInternoGCM")]
        public async Task<ActionResult> ActualizarKardexInternoGCM(int idKardex, string comentarios)
        {
            int idUsuario = Shared.ObtenerUsuarioSesion(HttpContext.User.Identity);
            ResponseModel<string> response  = await _controlCalidadServices.ActualizarKardexInternoGCM(idKardex, comentarios,idUsuario);
            return Ok(response);
        }


        [HttpGet("ExportarOrdenFabricacionCaja")]
        public async Task<ActionResult> ExportarOrdenFabricacionCaja()
        {
            ResponseModel<string> response = await _controlCalidadServices.ExportarOrdenFabricacionCaja();
            return Ok(response);
        }
    }
}
