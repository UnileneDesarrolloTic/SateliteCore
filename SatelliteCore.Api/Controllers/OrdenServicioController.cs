using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SatelliteCore.Api.CrossCutting.Helpers;
using SatelliteCore.Api.Models.Request.GestionOrdenesServicio;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Services.Contracts;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Controllers
{
    [Authorize]
    [Route("api/[controller]")]
    [ApiController]
    [Produces("application/json")]
    public class OrdenServicioController : ControllerBase
    {
        private readonly IOrdenServicioServices _ordenServicioService;

        public OrdenServicioController(IOrdenServicioServices ordenServicioService)
        {
            _ordenServicioService = ordenServicioService;
        }

        [HttpGet("listarOrdenServicio")]
        public async Task<IActionResult> ListarOrdenServicio(DateTime fechaInicio, DateTime fechaFin)
        {
            ResponseModel<IEnumerable<ListarOrdenServicioResponseDTO>> ordendes = await _ordenServicioService.ListarOrdenServicio(fechaInicio, fechaFin);
            return Ok(ordendes);
        }

        [HttpGet("listarTransportistaCombox")]
        public async Task<IActionResult> ListarTransportistaCombox()
        {
            ResponseModel<IEnumerable<ListaTransportistaComboxResponse>> ordendes = await _ordenServicioService.ListarTransportistaCombox();
            return Ok(ordendes);
        }

        [HttpGet("listaDetalleOrdenServicio")]
        public async Task<IActionResult> ListaDetalleOrdenServicio(int codigoOrdenServicio)
        {
            ResponseModel<IEnumerable<DetalleOrdenServicioResponse>> ordendes = await _ordenServicioService.ListaDetalleOrdenServicio(codigoOrdenServicio);
            return Ok(ordendes);
        }

        [HttpGet("listaGuiaRemision")]
        public async Task<IActionResult> ListaGuiaRemision(DateTime fechaInicio, DateTime fechaFin)
        {
            ResponseModel<IEnumerable<OrdenServicioGuiaRemisionResponse>> guias = await _ordenServicioService.ListaGuiaRemision(fechaInicio, fechaFin);
            return Ok(guias);
        }

        [HttpPost("modificarOrdenServicio")]
        public async Task<IActionResult> ModificarOrdenServicio(OrdenServicioModificadosDTO ordenes)
        {
            string usuarioSesion = Shared.ObtenerUsuarioSpring(HttpContext.User.Identity);
            ResponseModel<string> result = await _ordenServicioService.ModificarOrdenServicio(ordenes, usuarioSesion);
            return Ok(result);
        }

        [HttpGet("eliminarDetalle")]
        public async Task<IActionResult> EliminarDetalleOrdenServicio(int id)
        {
            ResponseModel<string> result = await _ordenServicioService.EliminarDetalleOrdenServicio(id);
            return Ok(result);
        }

        [HttpPost("editarGuiaRemision")]
        public async Task<IActionResult> EditarGuiaRemision(EditarGuiaOS_DTO datosGuia)
        {
            ResponseModel<string> result = await _ordenServicioService.EditarGuiaRemision(datosGuia);
            return Ok(result);
        }

        [HttpPost("guardarTransportista")]
        public async Task<IActionResult> GuardarTransportista(DatosTransportistaDTO datosTransportista)
        {
            ResponseModel<string> response = await _ordenServicioService.GuardarTransportista(datosTransportista);
            return Ok(response);
        }

        [HttpPost("nuevaOrdenServicio")]
        public async Task<IActionResult> NuevaOrdenServicio(DatosRegistrarOrdenServicioDTO ordenServicio)
        {
            ordenServicio.Usuario= Shared.ObtenerUsuarioSpring(HttpContext.User.Identity);
            ResponseModel<string> response = await _ordenServicioService.NuevaOrdenServicio(ordenServicio);
            return Ok(response);
        }

        [HttpGet("exportarOrdenServicioSalidas")]
        public async Task<IActionResult> ExportarOrdenServicioSalidas(DateTime? inicio, DateTime? fin)
        {
            ResponseModel<string> responseReporte = await _ordenServicioService.ExportarSalidas(inicio, fin);
            return Ok(responseReporte);
        }

        [HttpGet("exportarOrdenServicio")]
        public async Task<IActionResult> ExportarOrdenServicio(int id)
        {
            ResponseModel<string> responseReporte = await _ordenServicioService.ExportarOrdenServicio(id);
            return Ok(responseReporte);
        }


        [HttpGet("ordenServicioRetornada")]
        public async Task<IActionResult> OrdenServicioRetornada(string ordenServicio)
        {
            ResponseModel<DatosOServicioMarcadoDTO?> response = await _ordenServicioService.OrdenServicioRetornada(ordenServicio);
            return Ok(response);
        }


        [HttpGet("eliminarOrdenServicio")]
        public async Task<IActionResult> EliminarOrdenServicio(string ordenServicio)
        {
            ResponseModel<string> response = await _ordenServicioService.EliminarOrdenServicio(ordenServicio);
            return Ok(response);
        }

        [HttpGet("reporteGuiaOrdenServicio")]
        public async Task<IActionResult> ReporteGuiaOrdenServicio(DateTime? fechaInicio, DateTime? fechaFin)
        {
            ResponseModel<string> responseReporte = await _ordenServicioService.ReporteGuiaOrdenServicio(fechaInicio, fechaFin);
            return Ok(responseReporte);
        }
    }
}
