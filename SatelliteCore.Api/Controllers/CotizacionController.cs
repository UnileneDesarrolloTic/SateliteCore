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
    [Route("api/[controller]")]
    [ApiController]
    public class CotizacionController : ControllerBase
    {
        private readonly ICotizacionServices _cotizacionServices;

        public CotizacionController(ICotizacionServices cotizacionServices)
        {
            _cotizacionServices = cotizacionServices;
        }

        [HttpPost("Listar")]
        public async Task<IActionResult> Listar(DatosListarCotizacionesPaginado datos)
        {
            if (!ModelState.IsValid)
            {
                List<string> listaErrores = ModelState.Values.SelectMany(m => m.Errors).Select(e => e.ErrorMessage).ToList();
                throw new ValidationModelException(listaErrores);
            }

            PaginacionModel<CotizacionEntity> response = await _cotizacionServices.Listar(datos);
            return Ok(response);
        }

        [HttpGet("FormatoEstructura")]
        public async Task<IActionResult> FormatoEstructura(int codFormato)
        {
            if (codFormato == 0)
                throw new ValidationModelException("El formato es obligatorio");

            ObtenerEstructuraFormCotizacionModel estructura = await _cotizacionServices.FormatoEstructura(codFormato);
            return Ok(estructura);
        }

        [HttpGet("FormatoDatos")]
        public async Task<IActionResult> FormatoDatos(int idFormato, string cotizacion)
        {
            if (string.IsNullOrEmpty(cotizacion) || idFormato == 0)
                throw new ValidationModelException("El formato y el código de la cotización es obligatorio");

            (object cabecera, object detalle) estructura = await _cotizacionServices.FormatoDatos(idFormato, cotizacion);

            object result = new { estructura.cabecera, estructura.detalle };

            return Ok(result);
        }

        [HttpPost("Guardar")]
        public async Task<IActionResult> Guardar(ObtenerFormatoCotizacion reporte)
        {

            if (!ModelState.IsValid)
            {
                List<string> listaErrores = ModelState.Values.SelectMany(m => m.Errors).Select(e => e.ErrorMessage).ToList();
                throw new ValidationModelException(listaErrores);
            }

            if (reporte.IdFormato == 0)
                throw new ValidationModelException("El formato de la cotización es obligatorio");

            int usuarioSesion = Shared.ObtenerUsuarioSesion(HttpContext.User.Identity);

            ResponseModel<string> result = await _cotizacionServices.Guardar(reporte, usuarioSesion);

            return Ok(result);
        }

        [HttpGet("ObtenerReporte")]
        public async Task<IActionResult> ObtenerReporte(string codigoReporte)
        {
            if (string.IsNullOrEmpty(codigoReporte))
                throw new ValidationModelException("El código del documento es obligatorio");

            ResponseModel<string> documento = await _cotizacionServices.ObtenerReporte(codigoReporte);

            return Ok(documento);
        }

        [HttpGet("ObtenerDatosReporte")]
        public async Task<IActionResult> ObtenerDatosReporte(string codigoReporte)
        {
            if (string.IsNullOrEmpty(codigoReporte))
                throw new ValidationModelException("El código del reporte es obligatorio");

            var documento = await _cotizacionServices.ObtenerDatosReporte(codigoReporte);

            return Ok(documento);
        }

        [HttpGet("FormatosPorCliente")]
        public async Task<IActionResult> FormatosPorCliente(int idCliente)
        {
            IEnumerable<FormatosPorClienteModel> listaFormatos = await _cotizacionServices.FormatosPorCliente(idCliente);

            return Ok(listaFormatos);
        }

        [HttpGet("ReportesPorCotizacion")]
        public async Task<IActionResult> ReportesPorCotizacion(string cotizacion)
        {
            if (string.IsNullOrEmpty(cotizacion))
                throw new ValidationModelException("El código de la cotización es obligatorio");

            IEnumerable<ReportesGeneradosPorCotizacionModel> listaFormatos = await _cotizacionServices.ReportesPorCotizacion(cotizacion);

            return Ok(listaFormatos);
        }

        [HttpPut("Actualizar")]
        public async Task<IActionResult> Actualizar(ActualizarReporteCotizacionModel reporte)
        {
            if (!ModelState.IsValid)
            {
                List<string> listaErrores = ModelState.Values.SelectMany(m => m.Errors).Select(e => e.ErrorMessage).ToList();
                throw new ValidationModelException(listaErrores);
            }

            int usuarioSesion = Shared.ObtenerUsuarioSesion(HttpContext.User.Identity);

            await _cotizacionServices.Actualizar(reporte, usuarioSesion);

            return Ok();
        }
    }
}
