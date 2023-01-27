
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.CrossCutting.Helpers;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Request.Contabildad;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.Contabilidad;
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
    public class ContabilidadController : ControllerBase
    {
        private readonly IContabilidadService _ContabilidadService;

        public ContabilidadController(IContabilidadService ContabilidadService)
        {
            _ContabilidadService = ContabilidadService;

        }

        [HttpGet("ListarDetraccionContabilidad")]
        public async Task<ActionResult> ListarDetraccion()
        {
            IEnumerable<DetraccionesEntity> response = await _ContabilidadService.ListarDetraccion();
            return Ok(response);
        }


        [HttpPost("ProcesarDetraccionContabilidad")]
        public ActionResult ProcesarDetraccionContabilidad(DatosFormato64 dato)
        {
            int response = _ContabilidadService.ProcesarDetraccionContabilidad(dato);
            return Ok(response);
        }



        [HttpPost("GernerarBlogNotasDetraccion")]
        public ActionResult GenerarBlogNotasDetraccion(FormatoProcesoDetracciones dato)
        {
            string reporte = _ContabilidadService.GenerarBlogNotasDetraccion(dato);
            return Ok(reporte);
        }


        [HttpPost("ConsultarProductoCostoBase")]
        public async Task<ActionResult> ConsultarProductoCostoBase(DatosFormatoFiltrarAnalisisCostoRequest dato)
        {
            IEnumerable<DatosFormatoDatosProductoCostobase> listado = await _ContabilidadService.ConsultarProductoCostoBase(dato);
            return Ok(listado);
        }


        [HttpPost("ExportarExcelProductoCostoBase")]
        public async Task<ActionResult> ExportarExcelProductoCostoBase(DatosFormatoFiltrarAnalisisCostoRequest dato)
        {
            ResponseModel<string> listado = await _ContabilidadService.ExportarExcelProductoCostoBase(dato);
            return Ok(listado);
        }

        [HttpPost("ProcesarProductoExcel")]
        public async Task<ActionResult> ProcesarProductoExcel(DatosFormatoFiltrarAnalisisCostoRequest dato)
        {
            IEnumerable<DatosFormatoDatosProductoCostobase> listado = await _ContabilidadService.ProcesarProductoExcel(dato);
            return Ok(listado);
        }


        [HttpGet("ConsultarRecetaProducto")]
        public async Task<ActionResult> ConsultarRecetaProducto(string Item)
        {
            IEnumerable<DatosFormatoRecetaItemComponente> listado = await _ContabilidadService.ConsultarRecetaProducto(Item);
            return Ok(listado);
        }

        [HttpPost("ListarItemComponentePrecio")]
        public async Task<ActionResult> ListarItemComponentePrecio(DatosFormatosComponentPrecio dato)
        {
            IEnumerable<DatosFormatoComponentePrecioUnitario> listado = await _ContabilidadService.ListarItemComponentePrecio(dato);
            return Ok(listado);
        }

        [HttpPost("InformacionTransaccionKardex")]
        public async Task<ActionResult> InformacionTransaccionKardex(DatoFormatoFiltroTransaccionKardex dato)
        {
            if (!ModelState.IsValid)
                throw new ValidationModelException("Los datos enviados no son válidos");

            InformacionTransaccionKardex Informacion = await _ContabilidadService.InformacionTransaccionKardex(dato);
            return Ok(Informacion);
        }

        [HttpPost("RegistrarInformacionTransaccionKardex")]
        public async Task<ActionResult> RegistrarInformacionTransaccionKardex(DatoFormatoFiltroTransaccionKardex dato)
        {

            if (!ModelState.IsValid)
                throw new ValidationModelException("Los datos enviados no son válidos");

            string usuario = Shared.ObtenerUsuarioSpring(HttpContext.User.Identity);
            ResponseModel<string> resultado = await _ContabilidadService.RegistrarInformacionTransaccionKardex(dato, usuario);
            return Ok(resultado);

        }

        [HttpGet("ListarInformacionReporteCierrePeriodo")]
        public async Task<ActionResult> ListarInformacionReporteCierrePeriodo(string periodo)
        {
            if (string.IsNullOrEmpty(periodo))
                throw new ValidationModelException("El Periodo no válido.");

            ResponseModel<IEnumerable<FormatoDatosCierreHistorico>> Resultado = await _ContabilidadService.ListarInformacionReporteCierrePeriodo(periodo);
            return Ok(Resultado);
        }

        [HttpGet("ListarInformacionReporteAnio")]
        public async Task<ActionResult> ListarInformacionReporteCierreAnio(int anio)
        {
            if (string.IsNullOrEmpty(anio.ToString()))
                throw new ValidationModelException("El año no es valido");

            ResponseModel<IEnumerable<FormatoDatosCierreHistorico>> Resultado = await _ContabilidadService.ListarInformacionReporteCierreAnio(anio);
            return Ok(Resultado);

        }

        [HttpGet("ListarDetalleReporteCierre")]
        public async Task<ActionResult> ListarDetalleReporteCierre(int Id, string Periodo, string Tipo)
        {   
            if(string.IsNullOrEmpty(Periodo))
                throw new ValidationModelException("El Periodo no válido.");

            ResponseModel<IEnumerable<DatosFormatoMostrarDetalleReporte>> Resultado = await _ContabilidadService.ListarDetalleReporteCierre(Id, Periodo, Tipo);
            return Ok(Resultado);
        }

        [HttpGet("AnularReporteCierre")]
        public async Task<ActionResult> AnularReporteCierre(int Id)
        {   
            if (string.IsNullOrEmpty(Id.ToString()) || Id == 0)
                throw new ValidationModelException("El Dato enviado no es válido.");

            string usuario = Shared.ObtenerUsuarioSpring(HttpContext.User.Identity);
            ResponseModel<string> resultado = await _ContabilidadService.AnularReporteCierre(Id, usuario);
            return Ok(resultado);
        }

        [HttpPost("RestablecerReporteCierre")]
        public async Task<ActionResult> RestablecerReporteCierre(DatosFormatoRestablecerCierre dato)
        {
            string usuario = Shared.ObtenerUsuarioSpring(HttpContext.User.Identity);
            ResponseModel<string> resultado = await _ContabilidadService.RestablecerReporteCierre(dato, usuario);
            return Ok(resultado);
        }
    }
}