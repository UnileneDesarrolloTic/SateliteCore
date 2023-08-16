using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using SatelliteCore.Api.CrossCutting.Helpers;
using SatelliteCore.Api.Models.Encajado;
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
    public class EncajadoController : ControllerBase
    {
        private readonly IEncajadoServices _encajadoServices;

        public EncajadoController(IEncajadoServices encajadoServices)
        {
            _encajadoServices = encajadoServices;
        }

        [HttpGet("listaOrdenesFabricacion")]
        public async Task<IActionResult> ListaOrdenesFabricacion(string ordenFabricacion, string lote)
        {
            ResponseModel<List<ListaOrdenesFabricaciónDTO>> listaOrdenes = await _encajadoServices.ListaOrdenesFabricacion(ordenFabricacion, lote);
            return Ok(listaOrdenes);
        }

        [HttpGet("listaTransferenciasEncaje")]
        public async Task<IActionResult> ListaTransferenciasEncaje(string ordenFabricacion)
        {
            ResponseModel<List<TransferenciaEncajadoDTO>> lista = await _encajadoServices.ListaTransferenciasEncaje(ordenFabricacion);

            return Ok(lista);
        }

        [HttpGet("registarNuevaTrasnferencia")]
        public async Task<IActionResult> RegistarNuevaTrasnferencia(string ordenFabricacion, decimal cantidad)
        {
            string usuario = Shared.ObtenerUsuarioSpring(HttpContext.User.Identity);
            ResponseModel<string> registro = await _encajadoServices.RegistarNuevaTrasnferencia(ordenFabricacion, cantidad, usuario);

            return Ok(registro);
        }

        [HttpGet("listraAsignacionesEncajePorEtapa")]
        public async Task<IActionResult> ListraAsignacionesEncajePorEtapa(int idEncaje, int etapa)
        {
            ResponseModel<object> result = await _encajadoServices.ListraAsignacionesEncajePorEtapa(idEncaje, etapa);

            return Ok(result);
        }

        [HttpPost("registrarAsignacion")]
        public async Task<IActionResult> RegistrarAsignacion(DatosRegistrarAsignacionDTO asignacion)
        {
            asignacion.UsuarioRegistro = Shared.ObtenerUsuarioSpring(HttpContext.User.Identity);
            ResponseModel<string> registro = await _encajadoServices.RegistrarAsignacion(asignacion);
            return Ok(registro);
        }

        [HttpGet("actualizaEstadoAsignacion")]
        public async Task<IActionResult> ActualizaEstadoAsignacion(int id, string estado)
        {
            string usuario = Shared.ObtenerUsuarioSpring(HttpContext.User.Identity);
            ResponseModel<string> response = await _encajadoServices.ActualizaEstadoAsignacion(id, estado, usuario);

            return Ok(response);
        }

        [HttpGet("reporteAsignacion")]
        public async Task<IActionResult> ReporteAsignacion(DateTime fechaInicio, DateTime fechaFin)
        {
            ResponseModel<string> response = await _encajadoServices.ReporteAsignacion(fechaInicio, fechaFin);
            return Ok(response);
        }
    }
}
