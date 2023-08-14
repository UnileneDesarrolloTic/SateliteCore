using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SatelliteCore.Api.CrossCutting.Helpers;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Request.Dispensacion;
using SatelliteCore.Api.Models.Request.GestionGuias;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.Dispensacion;
using SatelliteCore.Api.Models.Response.Logistica;
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
    public class DispensacionController : ControllerBase
    {

        private readonly IDispensacionServices _dispensacionServices;

        public DispensacionController(IDispensacionServices dispensacionServices)
        {
            _dispensacionServices = dispensacionServices;
        }


        [HttpPost("ObtenerOrdenFabricacion")]
        public async Task<IActionResult> ObtenerOrdenFabricacion(DatosFormatoFiltroOrdenFabricacion dato)
        {
            ResponseModel<IEnumerable<DatosFormatoObtenerOrdenFabricacion>> listado = await _dispensacionServices.ObtenerOrdenFabricacion(dato);
            return Ok(listado);
        }

        [HttpGet("RecetasOrdenFabricacion")]
        public async Task<IActionResult> RecetasOrdenFabricacion(string ordenFabricacion)
        {
            IEnumerable<DatosFormatoListadoMateriaPrimaDispensacion> listado = await _dispensacionServices.RecetasOrdenFabricacion(ordenFabricacion);
            return Ok(listado);
        }

        [HttpPost("RegistrarDispensacionMP")]
        public async Task<IActionResult> RegistrarDispensacionMP(DatosFormatoDispensacionMateriaPrima dato)
        {
            string usuario = Shared.ObtenerUsuarioSpring(HttpContext.User.Identity);
            ResponseModel<string> respuesta = await _dispensacionServices.RegistrarDispensacionMP(dato, usuario);
            return Ok(respuesta);
        }

        [HttpGet("HistorialDispensacionMP")]
        public async Task<IActionResult> HistorialDispensacionMP(string ordenFabricacion, string lote)
        {
            IEnumerable<DatosFormatoHistorialDispensaccion> respuesta = await _dispensacionServices.HistorialDispensacionMP(ordenFabricacion, lote);
            return Ok(respuesta);
        }

        [HttpGet("InformacionItem")]
        public async Task<IActionResult> InformacionItem(string item, string ordenFabricacion, string secuencia)
        {
            ResponseModel<DatosFormatoInformacionDispensacionPT> clasificacionArea = await _dispensacionServices.InformacionItem(item, ordenFabricacion, secuencia);
            return Ok(clasificacionArea);
        }

        [HttpGet("DetalleDispensacionReceta")]
        public async Task<IActionResult> DetalleDispensacionReceta()
        {
            DatosFormatoDispensacionDetalle detalleRecetaDispensacion = await _dispensacionServices.DetalleDispensacionReceta();
            return Ok(detalleRecetaDispensacion);
        }

        [HttpPost("RegistrarRecetasGlobal")]
        public async Task<IActionResult> RegistrarRecetasGlobal(List<DatosFormatoRegistroDispensacionRecetaGlobal> dato)
        {
            string usuario = Shared.ObtenerUsuarioSpring(HttpContext.User.Identity);
            ResponseModel<string> respuesta = await _dispensacionServices.RegistrarRecetasGlobal(dato, usuario);
            return Ok(respuesta);
        }

        [HttpPost("DispensacionGuiaDespacho")]
        public async Task<IActionResult> DispensacionGuiaDespacho(DatosFormatoFiltroDispensacion dato)
        {
            IEnumerable<DatosFormatoDispensacionGuiaDespacho> listado = await _dispensacionServices.DispensacionGuiaDespacho(dato);
            return Ok(listado);
        }


        [HttpGet("MostrarDispensacionDespacho")]
        public async Task<IActionResult> MostrarDispensacionDespacho(string id)
        {
            IEnumerable<DatosFormatoMostrarDispensacionDespacho> listado = await _dispensacionServices.MostrarDispensacionDespacho(id);
            return Ok(listado);
        }

        [HttpGet("GeneracionPdfDespacho")]
        public ResponseModel<string> GeneracionPdfDespacho(string id)
        {
            ResponseModel<string> listado =  _dispensacionServices.GeneracionPdfDespacho(id);
            return listado;
        }


    }
}
