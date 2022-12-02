using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.CrossCutting.Helpers;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Services.Contracts;
using System.Collections.Generic;
using System.Threading.Tasks;
using SystemsIntegration.Api.Models.Exceptions;


namespace SatelliteCore.Api.Controllers
{
    [Authorize]
    [ApiController]
    [Route("api/[controller]")]
    [Produces("application/json")]
    public class ExportacionesController : ControllerBase
    {
        private readonly IExportacionesServices _exportacionesServices;

        public ExportacionesController(IExportacionesServices exportacionesServices)
        {
            _exportacionesServices = exportacionesServices;
        }

        [HttpPost("ListarCotizacionExportaciones")]
        public async Task<IActionResult> ListarCotizacionExportaciones(FiltrarCotizacionExportacionModel filtro)
        {

            IEnumerable<DatosFormatoListarCotizacionExportacion> listaCotizacionExportacion = await _exportacionesServices.ListarCotizacionExportaciones(filtro);

            return Ok(listaCotizacionExportacion);
        }

        [HttpGet("BuscarCotizacionExportaciones")]
        public async Task<IActionResult> BuscarCotizacionExportaciones(string NumeroDocumento)
        {
            (object cabecera, object detalle) = await _exportacionesServices.BuscarCotizacionExportaciones(NumeroDocumento);
            object response = new { cabecera, detalle };

            return Ok(response);
        }

        [HttpPost("GuardarCotizacionExportaciones")]
        public async Task<IActionResult> GuardarCotizacionExportaciones(DatosFormatoFormularioCotizacionExportaciones datos)
        {
            string UsuarioSesion = Shared.ObtenerUsuarioSpring(HttpContext.User.Identity);

            ResponseModel<string> listaCotizacionExportacion = await _exportacionesServices.GuardarCotizacionExportaciones(datos, UsuarioSesion);

            return Ok(listaCotizacionExportacion);
        }


        [HttpPost("ProcesarExcelExportaciones")]
        public async Task<IActionResult> ProcesarExcelExportaciones(DatosFormato64 dato)
        {
            ResponseModel<List<FormatoDetalleCotizacionExportaciones>> response = await _exportacionesServices.ProcesarExcelExportaciones(dato);
            return Ok(response);
        }

        [HttpGet("BuscarItemMast")]
        public async Task<IActionResult> BuscarWHItemMast(string Opcion, string Descripcion)
        {
            ResponseModel<List<FormatoDetalleCotizacionExportaciones>> response = await _exportacionesServices.BuscarWHItemMast(Opcion, Descripcion);
            return Ok(response);
        }

        [HttpGet("DesactivarItemCotizacionExportacion")]
        public async Task<IActionResult> DesactivarItemCotizacionExportacion(string NumeroDocumento, string Item , int Linea)
        {
            string UsuarioSesion = Shared.ObtenerUsuarioSpring(HttpContext.User.Identity);

            ResponseModel<string> response = await _exportacionesServices.DesactivarItemCotizacionExportacion(NumeroDocumento, Item, Linea , UsuarioSesion);
            return Ok(response);
        }
    }
}
