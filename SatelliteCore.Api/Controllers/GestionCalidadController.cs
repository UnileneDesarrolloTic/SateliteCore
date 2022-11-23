using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SatelliteCore.Api.CrossCutting.Helpers;
using SatelliteCore.Api.Models.Dto.GestionCalidad;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request.GestionCalidad;
using SatelliteCore.Api.Models.Request.GestorDocumentario;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.GestionCalidad;
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
    public class GestionCalidadController : ControllerBase
    {
        private readonly IGestionCalidadServices _gestionCalidadServices;

        public GestionCalidadController(IGestionCalidadServices gestionCalidadServices)
        {
            _gestionCalidadServices = gestionCalidadServices;
        }

        [HttpGet("ListarMateriaPrima")]
        public async Task<IActionResult> ListarMateriaPrima(string tipoConsulta, string lote)
        {
            List<MateriaPrimaDTO> listaAnalisis = await _gestionCalidadServices.ObtenerMateriaPrima(tipoConsulta, lote);
            return Ok(listaAnalisis);
        }

        [HttpPost("DetalleSeguimientoPorLote")]
        public async Task<IActionResult> DetalleSeguimientoPorLote(RequestLotesDetalleDTO filtros)
        {
            DetalleSeguimientoLoteDTO detalle = await _gestionCalidadServices.DetalleSeguimientoPorLote(filtros);
            return Ok(detalle);
        }

        [HttpPost("VentasPorCliente")]
        public async Task<IActionResult> VentasPorCliente(RequestFiltroVentaCliente filtros)
        {
            List<VentasPorClienteDTO> detalle = await _gestionCalidadServices.VentasPorCliente(filtros);
            return Ok(detalle);
        }

        [HttpPost("RptVentasPorCliente")]
        public async Task<IActionResult> ReporteVentasPorCliente(RequestFiltroVentaCliente filtros)
        {
            ResponseModel<string> reporte = await _gestionCalidadServices.ReporteVentasPorCliente(filtros);
            return Ok(reporte);
        }

        [HttpPost("ListaReclamos")]
        public async Task<IActionResult> ListarReclamosQuejas(FiltrosListaReclamosDTO filtros)
        {
            PaginacionModel<ListaReclamosDTO> listaReclamos = await _gestionCalidadServices.ListarReclamosQuejas(filtros);
            return Ok(listaReclamos);
        }

        [HttpGet("RegistrarReclamo")]
        public async Task<IActionResult> ListarReclamosQuejas(int codigoCliente)
        {
            string usuarioSesion = Shared.ObtenerUsuarioSpringSesion(HttpContext.User.Identity);

            ResponseModel<object> listaReclamos = await _gestionCalidadServices.RegistrarReclamoCabecera(codigoCliente, usuarioSesion);

            return Ok(listaReclamos);
        }

        [HttpGet("DetalleReclamo")]
        public async Task<IActionResult> ObtenerDetalleReclamo(string codigoReclamo)
        {
            string usuarioSesion = Shared.ObtenerUsuarioSpringSesion(HttpContext.User.Identity);

            ResponseModel<ReclamoDTO> listaReclamos = await _gestionCalidadServices.ObtenerDetalleReclamo(codigoReclamo);

            return Ok(listaReclamos);
        }

        [HttpPost("LotesFiltradosReclamo")]
        public async Task<IActionResult> LotesFiltradosReclamo(FiltrosLotesReclamosDTO filtros)
        {
            ResponseModel<IEnumerable<LotesFiltradosReclamo>> reporte = await _gestionCalidadServices.LotesFiltradosReclamo(filtros);
            return Ok(reporte);
        }

        [HttpGet("ObtenerDatosItemLote")]
        public async Task<IActionResult> ObtenerDatosItemLote(string lote)
        {
            ResponseModel<DatosLoteReclamoDTO> listaReclamos = await _gestionCalidadServices.DatosItemLote(lote);
            return Ok(listaReclamos);
        }

        [HttpPost("GuardarReclamoDetalle")]
        public async Task<IActionResult> GuardarReclamoDetalle(TBDReclamosEntity detalle)
        { 
            detalle.UsuarioRegistro = Shared.ObtenerUsuarioSpringSesion(HttpContext.User.Identity);
            ResponseModel<object> result = await _gestionCalidadServices.GuardarDetalleReclamo(detalle);
            return Ok(result);
        }

        [HttpPost("ActualizarDetalleLoteReclamo")]
        public async Task<IActionResult> ActualizarDetalleLoteReclamo(TBDReclamosEntity detalle)
        {
            detalle.UsuarioRegistro = Shared.ObtenerUsuarioSpringSesion(HttpContext.User.Identity);
            ResponseModel<string> result = await _gestionCalidadServices.ActualizarDetalleLoteReclamo(detalle);
            return Ok(result);
        }

        [HttpGet("DatosReclamoLote")]
        public async Task<IActionResult> DatosReclamoLote (string codReclamo, string lote, string documento)
        {
            ResponseModel<CabeceraReclamoLoteDTO> listaReclamos = await _gestionCalidadServices.DataReclamoLote(codReclamo, lote, documento);
            return Ok(listaReclamos);
        }

        [HttpPost("ResponderReclamo")]
        public async Task<IActionResult> RegistrarRespuestaReclamo(RespuestaReclamoDTO respuesta)
        {
            respuesta.Usuario = Shared.ObtenerUsuarioSpringSesion(HttpContext.User.Identity);

            ResponseModel<string> result = await _gestionCalidadServices.RegistrarRespuestaReclamo(respuesta);
            return Ok(result);
        }

    }
}
