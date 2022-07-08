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
using System.Threading.Tasks;
using SystemsIntegration.Api.Models.Exceptions;

namespace SatelliteCore.Api.Controllers
{
    [Authorize]
    [ApiController]
    [Route("api/[controller]")]
    [Produces("application/json")]
    public class ComercialController : ControllerBase
    {
        private readonly IComercialServices _comercialServices;
        private readonly IAppConfig _appConfig;

        public ComercialController(IComercialServices comercialServices, IAppConfig appConfig)
        {
            _comercialServices = comercialServices;
            _appConfig = appConfig;
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

            (List<CotizacionEntity> lista, int totalRegistros) certificados = await _comercialServices.ListarCotizaciones(datos);

            PaginacionModel<CotizacionEntity> response
                    = new PaginacionModel<CotizacionEntity>(certificados.lista, datos.Pagina, datos.RegistrosPorPagina, certificados.totalRegistros);

            return Ok(response);
        }

        [HttpPost("ObtenerEstructuraFormato")]
        public async Task<ActionResult> ObtenerEstructuraFormato(DatosEstructuraFormatoCotizacion datos)
        {
            FormatoCotizacionEntity estructura = await _comercialServices.ObtenerEstructuraFormato(datos);
            return Ok(estructura);
        }

        [HttpPost("GenerarReporteCotizacion")]
        public async Task<ActionResult> GenerarReporteCotizacion(DatosReporteCotizacion datos)
        {
            try
            {
                string Reporte = string.Empty;
                //Validación para elegir reporte
                switch (datos.IdFormato)
                {
                    case 1:
                        Reporte = "01_Instituto+Nac.+de+Salud+Niño+Sede+San+Borja&rs:Command=Render";
                        break;
                    case 3:
                        Reporte = "03_Essalud+Incor&rs:Command=Render";
                        break;
                    case 4:
                        Reporte = "04_Essalud+Apurimac&rs:Command=Render";
                        break;
                    case 5:
                        Reporte = "05_Essalud+Almenara";
                        break;
                    case 6:
                        Reporte = "06_Essalud+Cusco+Red+Asistencial&rs:Command=Render";
                        break;
                    case 7:
                        Reporte = "07_Hospital+San+Bartolome&rs:Command=Render";
                        break;
                    case 9:
                        Reporte = "09_Essalud Rebagliati&rs:Command=Render";
                        break;
                    case 10:
                        Reporte = "10_Essalud+Junin&rs:Command=Render";
                        break;
                    case 11:
                        Reporte = "11_Essalud+Arequipa&rs:Command=Render";
                        break;
                    case 13:
                        Reporte = "13_Essalud Sabogal&rs:Command=Render";
                        break;
                    default:
                        Reporte = "12_FormatoGeneral&rs:Command=Render";
                        break;
                }
                string Formato = "&rs:Format=excel";
                string Parametros = "&NumeroDocumento=" + datos.NumeroDocumento;

                var theURL = _appConfig.ReportComercialFormatoCotizacion + Reporte + Parametros + Formato;

                var httpClientHandler = new HttpClientHandler()
                {
                    UseDefaultCredentials = true
                };

                HttpClient webClient = new HttpClient(httpClientHandler);

                Byte[] result = await webClient.GetByteArrayAsync(theURL);
                string base64String = Convert.ToBase64String(result, 0, result.Length);
                ResponseModel<string> response
                        = new ResponseModel<string>(true, "El reporte se generó correctamente", base64String);
                return Ok(response);
            }
            catch (Exception ex)
            {
                ResponseModel<string> response
                        = new ResponseModel<string>(true, "El reporte no se generó", ex.Message);
                return BadRequest(response);
            }

        }

        [HttpPost("RegistrarRespuestas")]
        public async Task<ActionResult> RegistrarRespuestas(FormatoCotizacionRespuesta datos)
        {
            int respuesta = await _comercialServices.RegistrarRespuestas(datos);
            return Ok(respuesta);
        }

        [HttpPost("GenerarReporteProtocoloAnalisis")]
        public async Task<ActionResult> GenerarReporteProtocoloAnalisis(DatosReporteProtocoloAnalisis datos)
        {
            try
            {
                string Reporte = "ProtocoloAnalisis&rs:Command=Render";
                string Formato = "&rs:Format=pdf";
                string Parametros = "&Lote=" + datos.Lote;

                var theURL = _appConfig.ReportComercialProtocoloAnalisis + Reporte + Parametros + Formato;

                var httpClientHandler = new HttpClientHandler()
                {
                    UseDefaultCredentials = true
                };

                HttpClient webClient = new HttpClient(httpClientHandler);

                Byte[] result = await webClient.GetByteArrayAsync(theURL);

                string base64String = Convert.ToBase64String(result, 0, result.Length);
                ResponseModel<string> response
                        = new ResponseModel<string>(true, "El reporte se generó correctamente", base64String);
                return Ok(response);
            }
            catch (Exception ex)
            {
                ResponseModel<string> response
                        = new ResponseModel<string>(false, "El reporte no se generó", ex.Message);
                return BadRequest(response);
            }

        }

        [HttpPost("ListarProtocoloAnalisis")]
        public async Task<ActionResult> ListarProtocoloAnalisis(DatosProtocoloAnalisisListado datos)
        {
            try
            {
                (List<DetalleProtocoloAnalisis> lista, int totalRegistros) result = await _comercialServices.ListarProtocoloAnalisis(datos);

                PaginacionModel<DetalleProtocoloAnalisis> response
                        = new PaginacionModel<DetalleProtocoloAnalisis>(result.lista, datos.Pagina, datos.RegistrosPorPagina, result.totalRegistros);

                return Ok(response);
            }
            catch (Exception ex)
            {
                ResponseModel<string> response
                        = new ResponseModel<string>(false, "La lista no se pudo cargar", ex.Message);
                return BadRequest(response);
            }
        }

        [HttpPost("ListarClientes")]
        public async Task<ActionResult> ListarClientes()
        {
            try
            {
                List<DetalleClientes> result = await _comercialServices.ListarClientes();
                ResponseModel<List<DetalleClientes>> response
                        = new ResponseModel<List<DetalleClientes>>(true, "La lista se cargó correctamente", result);

                return Ok(response);
            }
            catch (Exception ex)
            {
                ResponseModel<string> response
                        = new ResponseModel<string>(false, "La lista no se pudo cargar", ex.Message);
                return BadRequest(response);
            }

        }

        [HttpPost("ListarDocumentoLicitacion")]
        public async Task<ActionResult> ListarDocumentoLicitacion(DatosFormatoDocumentoLicitacion datos)
        {   
            IEnumerable<FormatoLicitaciones> response = await _comercialServices.ListarDocumentoLicitacion(datos);
            return Ok(response);
        }

       [HttpPost("NumerodeGuiaLicitacion")]
       public async Task<ActionResult> NumerodeGuiaLicitacion(List<FormatoLicitacionesOT> datos)
        {
          
            try
            {
                if (datos.Count == 0)
                {
                    throw new ValidationModelException("Debe Seleccionar una o varias guias");
                  
                }

                ResponseModel<string> response = await _comercialServices.NumerodeGuiaLicitacion(datos);
                return Ok(response);
            }
            catch (Exception ex)
            {
                ResponseModel<string> response = new ResponseModel<string>(false, "La lista no se pudo cargar", ex.Message);
                return BadRequest(response);
            }
        }

       [HttpGet("NumeroPedido")]
        public async Task<IActionResult> NumeroPedido(string pedido)
        {
            if(string.IsNullOrEmpty(pedido))
                throw new ValidationModelException("Los datos enviados no son válidos");
            

            DatoPedidoDocumentoModel  NumeroPedido = await _comercialServices.NumeroPedido(pedido);

            if (NumeroPedido.NumeroDocumento == null) {
                ResponseModel<DatoPedidoDocumentoModel> response = new ResponseModel<DatoPedidoDocumentoModel>(false, "No Existe ese Numero Pedido", NumeroPedido);
                return Ok(response);
            }
            else
            {
                ResponseModel<DatoPedidoDocumentoModel> response = new ResponseModel<DatoPedidoDocumentoModel>(true, Constante.MESSAGE_SUCCESS, NumeroPedido);
                return Ok(response);
            }
            
        }

        [HttpPost("RegistrarRotuladosPedido")]
        public async Task<ActionResult> RegistrarRotuladosPedido(DatosEstructuraNumeroRotuloModel dato)
        {
            int idUsuario = Shared.ObtenerUsuarioSesion(HttpContext.User.Identity);
            ResponseModel<string> response = await _comercialServices.RegistrarRotuladosPedido(dato, idUsuario);

            return Ok(response);
        }


        [HttpPost("ListarGuiaporFacturar")]
        public async Task<ActionResult> ListarGuiaporFacturar(DatosEstructuraGuiaPorFacturarModel dato)
        {
            IEnumerable<FormatoGuiaPorFacturarModel> listar = await _comercialServices.ListarGuiaporFacturar(dato);
            return Ok(listar);
        }


        [HttpPost("RegistrarGuiaporFacturar")]
        public async Task<ActionResult> RegistrarGuiaporFacturar(DatoFormatoEstructuraGuiaFacturada dato)
        {
            int idUsuario = Shared.ObtenerUsuarioSesion(HttpContext.User.Identity);

            await _comercialServices.RegistrarGuiaporFacturar(dato, idUsuario);

            return Ok();
        }



        [HttpPost("ListarGuiaporFacturarExportar")]
        public async Task<ActionResult> ListarGuiaporFacturarExportar(DatosEstructuraGuiaPorFacturarModel dato)
        {   

            string listar = await _comercialServices.ListarGuiaporFacturarExportar(dato);
            return Ok(listar);
        }

    }
}