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
using System.Security.Claims;
using System.Threading.Tasks;
using SystemsIntegration.Api.Models.Exceptions;

namespace SatelliteCore.Api.Controllers
{
    [Authorize]
    [ApiController]
    [Route("api/[controller]")]
    [Produces("application/json")]
    public class ControlCalidadController : ControllerBase
    {
        private readonly IControlCalidadServices _controlCalidadServices;
        private readonly IAppConfig _appConfig;
        public ControlCalidadController(IControlCalidadServices controlCalidadServices, IAppConfig appConfig)
        {
            _controlCalidadServices = controlCalidadServices;
            _appConfig = appConfig;
        }

        [HttpPost("ListarCertificados")]
        public async Task<ActionResult> ListarCertificados(DatosListarCertificadoPaginado datos)
        {
            try
            {
                if (!ModelState.IsValid)
                {
                    ResponseModel<string> responseError =
                            new ResponseModel<string>(false, Constante.MODEL_VALIDATION_FAILED, "");

                    return BadRequest(responseError);
                }

                (List<CertificadoEsterilizacionEntity> lista, int totalRegistros) = await _controlCalidadServices.ListarCertificados(datos);

                PaginacionModel<CertificadoEsterilizacionEntity> response
                        = new PaginacionModel<CertificadoEsterilizacionEntity>(lista, datos.Pagina, datos.RegistrosPorPagina, totalRegistros);
                return Ok(response);
            }
            catch (Exception ex)
            {

                return BadRequest(ex.Message);
            }

        }

        [HttpPost("RegistrarCertificado")]
        public ActionResult RegistrarCertificado(CertificadoEsterilizacionEntity certificado)
        {
            try
            {
                _controlCalidadServices.RegistrarCertificado(certificado);

                return Ok();
            }
            catch (Exception ex)
            {
                ResponseModel<string> response
                        = new ResponseModel<string>(true, "No se completó ", ex.Message);
                return BadRequest(response);
            }


        }

        [HttpPost("ListarLotes")]
        public async Task<ActionResult> ListarLotes(DatosLote datos)
        {
            if (!ModelState.IsValid)
            {
                ResponseModel<string> responseError =
                        new ResponseModel<string>(false, Constante.MODEL_VALIDATION_FAILED, "");

                return BadRequest(responseError);
            }

            (List<LoteEntity> lista, int totalRegistros) = await _controlCalidadServices.ListarLotes(datos);

            PaginacionModel<LoteEntity> response
                    = new PaginacionModel<LoteEntity>(lista, datos.Pagina, datos.RegistrosPorPagina, totalRegistros);

            return Ok(response);
        }

        [HttpPost("GenerarReporte")]
        public async Task<ActionResult> GenerarReporte(DatosReporte datos)
        {
            try
            {
                string ReporteEsterilizacion = "Certificado_Esterilizacion&rs:Command=Render";
                string Formato = "&rs:Format=PDF";
                string Parametros = "&Id=" + datos.Id;

                var theURL = _appConfig.ReportControlDeCalidad + ReporteEsterilizacion + Parametros + Formato;

                var httpClientHandler = new HttpClientHandler()
                {
                    UseDefaultCredentials = true
                };

                HttpClient webClient = new HttpClient(httpClientHandler);

                Byte[] result = await webClient.GetByteArrayAsync(theURL);
                string base64String = Convert.ToBase64String(result, 0, result.Length);
                ResponseModel<string> response = new ResponseModel<string>(true, "El reporte se generó correctamente", base64String);
                return Ok(response);
            }
            catch (Exception ex)
            {
                ResponseModel<string> response
                        = new ResponseModel<string>(true, "El reporte no se generó", ex.Message);
                return BadRequest(response);
            }

        }

        [HttpPost("RegistrarLote")]
        public async Task<ActionResult> RegistrarLote(LoteEntity lote)
        {
            int result = await _controlCalidadServices.RegistrarLote(lote);

            ResponseModel<int> response
                    = new ResponseModel<int>(true, "El lote se registró correctamente", result);

            return Ok(response);

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

            (List<CotizacionEntity> lista, int totalRegistros) = await _controlCalidadServices.ListarCotizaciones(datos);

            PaginacionModel<CotizacionEntity> response
                    = new PaginacionModel<CotizacionEntity>(lista, datos.Pagina, datos.RegistrosPorPagina, totalRegistros);

            return Ok(response);
        }

        [HttpGet("ObtenerInformacionLote")]
        public async Task<ActionResult> ObtenerInformacionLote(string NumeroLote)
        {
            if (NumeroLote == "")
            {
                throw new ValidationModelException("Debe Ingresar el Numero de Lote");
            }

            DatosFormatoListarOrdenFabricacionModel response =  await _controlCalidadServices.ObtenerInformacionLote(NumeroLote);
            return Ok(response);
        }


        [HttpGet("ListarTransaccionItem")]
        public async Task<ActionResult> ListarTransaccionItem(string NumeroLote, string codAlmacen)
        {
            if (NumeroLote == "")
            {
                ResponseModel<string> responseError = new ResponseModel<string>(false, Constante.MODEL_VALIDATION_FAILED, "");
                return BadRequest(responseError);
            }

            IEnumerable<DatosFormatoListarTransaccion> response = await _controlCalidadServices.ListarTransaccionItem(NumeroLote, codAlmacen);
            return Ok(response);
        }


        [HttpPost("RegistrarLoteNumeroCaja")]
        public async Task<ActionResult> RegistrarLoteNumeroCaja(DatosFormatoOrdenFabricacionRequest dato)
        {
            int idUsuario = Shared.ObtenerUsuarioSesion(HttpContext.User.Identity);
            ResponseModel<string> response = await _controlCalidadServices.RegistrarLoteNumeroCaja(dato, idUsuario);
            return Ok(response);
        }

        [HttpGet("ListarKardexInternoNumeroLote")]
        public async Task<ActionResult> ListarKardexInternoNumeroLote(string NumeroLote)
        {
            IEnumerable < DatosFormatoKardexInternoGCM > response= await _controlCalidadServices.ListarKardexInternoNumeroLote(NumeroLote);
            return Ok(response);
        }

        [HttpPost("RegistrarKardexInternoGCM")]
        public async Task<ActionResult> RegistrarKardexInternoGCM(DatosFormatoRegistrarKardexInternoGCM dato)
        {
            int idUsuario = Shared.ObtenerUsuarioSesion(HttpContext.User.Identity);
            ResponseModel<string> response = await _controlCalidadServices.RegistrarKardexInternoGCM(dato, idUsuario);
            return Ok(response);
        }

        [HttpGet("ActualizarKardexInternoGCM")]
        public async Task<ActionResult> ActualizarKardexInternoGCM(int idKardex, string comentarios)
        {
            int idUsuario = Shared.ObtenerUsuarioSesion(HttpContext.User.Identity);
            ResponseModel<string> response  = await _controlCalidadServices.ActualizarKardexInternoGCM(idKardex, comentarios,idUsuario);
            return Ok(response);
        }


        [HttpGet("ExportarOrdenFabricacionCaja")]
        public async Task<ActionResult> ExportarOrdenFabricacionCaja(string anioProduccion)
        {
            ResponseModel<string> response = await _controlCalidadServices.ExportarOrdenFabricacionCaja(anioProduccion);
            return Ok(response);
        }


        [HttpPost("ListarControlLotes")]
        public async Task<ActionResult> ListarControlLotes(DatosFormatoFiltrarControlLotesModel dato)
        {
            IEnumerable<DatosFormatosListarControlLotes> response = await _controlCalidadServices.ListarControlLotes(dato);
            return Ok(response);
        }

        [HttpPost("ActualizarControlLotes")]
        public async Task<ActionResult> ActualizarControlLotes(DatosFormatoControlLotesActualizarFEntrega dato)
        {
            ResponseModel<string> response = await _controlCalidadServices.ActualizarControlLotes(dato);
            return Ok(response);
        }


        [HttpGet("ListarMaestroNumeroParte")]
        public async Task<IActionResult> ListarMaestroNumeroParte(string Grupo,string Tabla)
        {
            IEnumerable<DatosFormatoTablaNumerodeParte> response = await _controlCalidadServices.ListarMaestroNumeroParte(Grupo, Tabla);
            return Ok(response);
        }

        [HttpGet("ListarAtributos")]
        public async Task<IActionResult> ListarAtributos()
        {
            IEnumerable<DatosFormatoTablaAbributoModel> response = await _controlCalidadServices.ListarAtributos();
            return Ok(response);
        }

        [HttpGet("ListarDescripcion")]
        public async Task<IActionResult> ListarDescripcion(string Marca, string Hebra)
        {
            IEnumerable<DatosFormatoTablaDescripcionModel> response = await _controlCalidadServices.ListarDescripcion(Marca, Hebra);
            return Ok(response);
        }

        [HttpGet("ListarLeyenda")]
        public async Task<IActionResult> ListarLeyenda(string Marca, string Hebra)
        {
            IEnumerable<DatosFormatoTablaLeyendaModel> response = await _controlCalidadServices.ListarLeyenda(Marca, Hebra);
            return Ok(response);
        }

        [HttpGet("ListarTablaPrueba")]
        public async Task<IActionResult> ListarTablaPrueba(string Metodologia)
        {
            IEnumerable<DatosFormatoTablaPruebasModel> response = await _controlCalidadServices.ListarTablaPrueba(Metodologia);
            return Ok(response);
        }

       

        [HttpGet("ListarObtenerAgujasDescripcionNuevo")]
        public async Task<IActionResult> ListarObtenerAgujasDescripcionNuevo()
        {
            IEnumerable<DatosFormatoObtenerTablaAgujasNuevoModel> response = await _controlCalidadServices.ListarObtenerAgujasDescripcionNuevo();
            return Ok(response);
        }

        [HttpPost("NuevoDescripcionDT")]
        public async Task<IActionResult> NuevoDescripcionDT(DatosFormatoActualizacionDescripcionModel dato)
        {
            string idUsuario = Shared.ObtenerUsuarioSesion(HttpContext.User.Identity).ToString();
            ResponseModel<string> response = await _controlCalidadServices.NuevoDescripcionDT(dato, idUsuario);
            return Ok(response);
        }

        [HttpGet("EliminarDescripcionDT")]
        public async Task<IActionResult> EliminarDescripcionDT(string IdDescripcion)
        {
            ResponseModel<string> response = await _controlCalidadServices.EliminarDescripcionDT(IdDescripcion);
            return Ok(response);
        }

        [HttpGet("ListarObtenerAgujasDescripcionActualizar")]
        public async Task<IActionResult> ListarObtenerAgujasDescripcionActualizar(string IdDescripcion)
        {
            IEnumerable<DatosFormatoObtenerAgujasDescripcionModel> response = await _controlCalidadServices.ListarObtenerAgujasDescripcionActualizar(IdDescripcion);
            return Ok(response);
        }

        [HttpPost("ActualizarDescripcionDT")]
        public async Task<IActionResult> ActualizarDescripcionDT(DatosFormatoActualizacionDescripcionModel dato)
        {
            string idUsuario = Shared.ObtenerUsuarioSesion(HttpContext.User.Identity).ToString();
            ResponseModel<string> response = await _controlCalidadServices.ActualizarDescripcionDT(dato, idUsuario);
            return Ok(response);
        }


        [HttpPost("RegistrarActualizarLeyendaDT")]
        public async Task<IActionResult> RegistrarActualizarLeyendaDT(DatosFormatoLeyendaDTModel dato)
        {
            string idUsuario = Shared.ObtenerUsuarioSesion(HttpContext.User.Identity).ToString();
            ResponseModel<string> response = await _controlCalidadServices.RegistrarActualizarLeyendaDT(dato, idUsuario);
            return Ok(response);
        }

        [HttpGet("EliminarLeyendaDT")]
        public async Task<IActionResult> EliminarLeyendaDT(string IdLeyenda)
        {
            ResponseModel<string> response = await _controlCalidadServices.EliminarLeyendaDT(IdLeyenda);
            return Ok(response);
        }


        [HttpPost("RegistrarActualizarPruebaDT")]
        public async Task<IActionResult> RegistrarActualizarPruebaDT(DatosFormatoNuevoPruebaModel dato)
        {
            string idUsuario = Shared.ObtenerUsuarioSesion(HttpContext.User.Identity).ToString();
            ResponseModel<string> response = await _controlCalidadServices.RegistrarActualizarPruebaDT(dato, idUsuario);
            return Ok(response);
        }


        [HttpGet("EliminarPruebaDT")]
        public async Task<IActionResult> EliminarPruebaDT(string IdPrueba)
        {
            ResponseModel<string> response = await _controlCalidadServices.EliminarPruebaDT(IdPrueba);
            return Ok(response);
        }


        //FORMATO DE PROTOCOLO

        [HttpGet("BuscarNumeroLoteProtocolo")]
        public async Task<IActionResult> BuscarNumeroLoteProtocolo(string NumeroLote)
        {
            if (NumeroLote == "")
            {
                ResponseModel<string> responseError = new ResponseModel<string>(false, Constante.MODEL_VALIDATION_FAILED, "");
                return BadRequest(responseError);
            }

            ResponseModel<DatosFormatoNumeroLoteProtocoloModel> response = await _controlCalidadServices.BuscarNumeroLoteProtocolo(NumeroLote);
            return Ok(response);
        }

        [HttpGet("BuscarPruebaFormatoProtocolo")]
        public async Task<IActionResult> BuscarPruebaFormatoProtocolo(string NumeroLote,string NumeroParte)
        {
            if (NumeroLote == "")
            {
                ResponseModel<string> responseError = new ResponseModel<string>(false, Constante.MODEL_VALIDATION_FAILED, "");
                return BadRequest(responseError);
            }

            IEnumerable<DatosFormatosDatoListarPruebaProtocolo> response = await _controlCalidadServices.BuscarPruebaFormatoProtocolo(NumeroLote, NumeroParte);
            return Ok(response);
        }

        [HttpPost("RegistrarControlProcesoProtocolo")]
        public async Task<IActionResult> RegistrarControlProcesoProtocolo(DatosFormatoControlProcesosProtocoloModel dato )
        {
            string idUsuario = Shared.ObtenerUsuarioSesion(HttpContext.User.Identity).ToString();
            ResponseModel<string> response = await _controlCalidadServices.RegistrarControlProcesoProtocolo(dato, idUsuario);
            return Ok(response);
        }

        [HttpPost("RegistrarControlPTProtocolo")]
        public async Task<IActionResult> RegistrarControlPTProtocolo(DatosFormatoControlProductoTermino dato)
        {
            string idUsuario = Shared.ObtenerUsuarioSesion(HttpContext.User.Identity).ToString();
            ResponseModel<string> response = await _controlCalidadServices.RegistrarControlPTProtocolo(dato, idUsuario);
            return Ok(response);
        }

        [HttpGet("RegistrarPruebasEfectuadasProtocolo")]
        public async Task<IActionResult> RegistrarPruebasEfectuadasProtocolo(DatosFormatoPruebasEfectuasProtocolos dato)
        {
            string idUsuario = Shared.ObtenerUsuarioSesion(HttpContext.User.Identity).ToString();
            ResponseModel<string> response = await _controlCalidadServices.RegistrarPruebasEfectuadasProtocolo(dato, idUsuario);
            return Ok(response);
        }

        [HttpGet("BuscarInformacionResultadoProtocolo")]
        public async Task<IActionResult> BuscarInformacionResultadoProtocolo(string NumeroLote)
        {

            IEnumerable<DatosFormatoInformacionResultadoProtocolo> response = await _controlCalidadServices.BuscarInformacionResultadoProtocolo(NumeroLote);
            return Ok(response);
        }

        [HttpPost("RegistrarFormatoProtocolo")]
        public async Task<IActionResult> InsertarCabeceraFormatoProtocolo(DatosFormatoCabeceraFormatoProtocolo dato)
        {
            string UsuarioSesion = Shared.ObtenerUsuarioSpring(HttpContext.User.Identity);

            ResponseModel<string> response = await _controlCalidadServices.InsertarCabeceraFormatoProtocolo(dato, UsuarioSesion);
            return Ok(response);
        }

        [HttpGet("ImprimirControlProcesoInterno")]
        public async Task<IActionResult> ImprimirControlProcesoInterno(string NumeroLote)
        {
            ResponseModel<string> response = await _controlCalidadServices.ImprimirControlProcesoInterno(NumeroLote);
            return Ok(response);
        }

        [HttpGet("ImprimirControlPruebas")]
        public async Task<IActionResult> ImprimirControlPruebas(string NumeroLote)
        {
            ResponseModel<string> response = await _controlCalidadServices.ImprimirControlPruebas(NumeroLote);
            return Ok(response);
        }


        [HttpGet("ImprimirDocumentoProtocolo")]
        public async Task<IActionResult> ImprimirDocumentoProtocolo(string NumeroLote, bool Opcion)
        {
            ResponseModel<string> response = await _controlCalidadServices.ImprimirDocumentoProtocolo(NumeroLote, Opcion);
            return Ok(response);
        }
    }
}
