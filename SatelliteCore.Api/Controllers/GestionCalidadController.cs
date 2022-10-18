using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SatelliteCore.Api.Models.Dto.GestionCalidad;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Services.Contracts;
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
    }
}
