using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SatelliteCore.Api.CrossCutting.Helpers;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Request.GestionGuias;
using SatelliteCore.Api.Models.Response;
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
    public class LogisticaController : ControllerBase
    {
        private readonly ILogisticaServices _logisticaServices;

        public LogisticaController(ILogisticaServices logisticaServices)
        {
            _logisticaServices = logisticaServices;
        }

        [HttpGet("ObtenerNumeroGuias")]
        public async Task<IActionResult> ObtenerNumeroGuias(string numeroguia)
        {
            string Usuario = Shared.ObtenerUsuarioSpring(HttpContext.User.Identity);
            ResponseModel<DatosFormatoPlanOrdenServicosD> respuesta = await _logisticaServices.ObtenerNumeroGuias(numeroguia, Usuario);
            return Ok(respuesta);
        }

        [HttpPost("RegistrarRetornoGuia")]
        public async Task<IActionResult> RegistrarRetornoGuia(DatosFormatoRetornoGuiaRequest dato)
        {
            if (!ModelState.IsValid)
                throw new ValidationModelException("Los datos de prueba de flexion no son válidos !!");

            string Usuario = Shared.ObtenerUsuarioSpring(HttpContext.User.Identity);
            ResponseModel<string> response = await _logisticaServices.RegistrarRetornoGuia(dato, Usuario);

            return Ok(response);
        }

        [HttpPost("ListarRetornoGuia")]
        public async Task<IActionResult> ListarRetornoGuia(DatosFormatoGestionGuiasClienteModel dato)
        {
            ResponseModel<IEnumerable<DatosFormatosReporteRetornoGuia>> response = await _logisticaServices.ListarRetornoGuia(dato);
            return Ok(response);
        }


        [HttpPost("ExportarExcelRetornoGuia")]
        public async Task<IActionResult> ExportarExcelRetornoGuia(DatosFormatoGestionGuiasClienteModel dato)
        {
            ResponseModel<string> response = await _logisticaServices.ExportarExcelRetornoGuia(dato);
            return Ok(response);
        }


        [HttpPost("ListarItemVentas")]
        public async Task<IActionResult> ListarItemVentas(FormatoDatosBusquedaItemsVentas dato)
        {
            IEnumerable<DatosFormatoItemVentas> result = await _logisticaServices.ListarItemVentas(dato);
            return Ok(result);
        }


        [HttpGet("BuscarItemVentas")]
        public async Task<IActionResult> BuscarItemVentas(string Item)
        {
            IEnumerable<DatosFormatoItemLoteAlmacen> result = await _logisticaServices.BuscarItemVentas(Item);
            return Ok(result);
        }

        [HttpPost("ListarItemVentasExportar")]
        public async Task<IActionResult> ListarItemVentasExportar(FormatoDatosBusquedaItemsVentas dato)
        {
            ResponseModel<string> response = await _logisticaServices.ListarItemVentasExportar(dato);

            return Ok(response);
        }

        [HttpGet("ListarItemVentasDetalle")]
        public async Task<IActionResult> ListarItemVentasDetalle()
        {
            IEnumerable<DatosFormatoDetalledelItemVentas> result = await _logisticaServices.ListarItemVentasDetalle();
            return Ok(result);
        }

        [HttpGet("ListarItemVentasDetalleExportar")]
        public async Task<IActionResult> ListarItemVentasDetalleExportar()
        {
            ResponseModel<string> response = await _logisticaServices.ListarItemVentasDetalleExportar();
            return Ok(response);
        }


        [HttpPost("DetalleComprometidoItem")]
        public async Task<IActionResult> DetalleComprometidoItem(DatosFormatoRequestDetalleComprometido dato)
        {
            IEnumerable<DatosFormatoDetalleComprometidoItem> result = await _logisticaServices.DetalleComprometidoItem(dato);
            return Ok(result);
        }

        [HttpGet("BuscarNumeroPedido")]
        public async Task<IActionResult> BuscarNumeroPedido(string NumeroDocumento, string Tipo)
        {
            IEnumerable<DatosFormatoMateriaPrimaItemLogistica> result = await _logisticaServices.BuscarNumeroPedido(NumeroDocumento, Tipo);
            return Ok(result);
        }

        [HttpGet("BuscardDetalleRecetaMP")]
        public async Task<IActionResult> BuscardDetalleRecetaMP(string Item, string Cantidad)
        {
            IEnumerable<DatosFormatoDetalleRecetaMPLogistica> result = await _logisticaServices.BuscardDetalleRecetaMP(Item, Cantidad);
            return Ok(result);
        }
    }
}
